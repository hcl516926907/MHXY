"""
套利计算模块

职责：
  - 读取历史 CSV 数据并聚合市场快照
  - 计算各商品 (买入天数, 卖出天数) 组合的利润与收益率
  - 导出每日精简 xlsx 报表（top_opportunities.xlsx）
"""

import os
import glob

import pandas as pd

from ocr_corrections import normalize_ocr_text
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from paths import DATA_ROOT, ANALYSIS_ROOT

# ── 套利规则常量 ───────────────────────────────────────────────────────────────
TAX_RATE = 0.09

# 冻结天数（按商品分类）；不在表中的默认 7 天
FREEZE_DAYS_BY_CATEGORY: dict[str, int] = {
    "高级兽诀": 30,
    "高级内丹": 30,
}
DEFAULT_FREEZE_DAYS = 7
# ─────────────────────────────────────────────────────────────────────────────


def load_all_source_csvs(up_to_date: str) -> pd.DataFrame:
    """
    读取 DATA_ROOT 下所有日期目录（≤ up_to_date）中的全部 CSV 文件。
    up_to_date 为 YYYYMMDD 格式字符串。
    """
    import re
    if not os.path.isdir(DATA_ROOT):
        raise FileNotFoundError(f"未找到数据根目录 data/，请确认路径配置正确。")

    date_dirs = sorted(
        d for d in os.listdir(DATA_ROOT)
        if re.fullmatch(r"\d{8}", d) and d <= up_to_date
    )
    if not date_dirs:
        raise FileNotFoundError(
            f"data/ 下没有日期 ≤ {up_to_date} 的目录，请确认已扫描过数据。"
        )

    frames = []
    for d in date_dirs:
        dir_path = os.path.join(DATA_ROOT, d)
        for path in glob.glob(os.path.join(dir_path, "*.csv")):
            try:
                frames.append(pd.read_csv(path, encoding="utf-8-sig"))
            except Exception as e:
                print(f"  [警告] 跳过文件 {d}/{os.path.basename(path)}：{e}")

    if not frames:
        raise ValueError("所有 CSV 文件均读取失败。")

    df = pd.concat(frames, ignore_index=True)
    # 修正历史 CSV 中可能存在的 OCR 误识别字符（如繁体字形 強→强）
    if "item_name" in df.columns:
        df["item_name"] = df["item_name"].astype(str).map(normalize_ocr_text)
    print(f"  已读取 {len(date_dirs)} 个日期目录（{date_dirs[0]} ~ {date_dirs[-1]}）")
    return df


def aggregate_market(df: pd.DataFrame) -> pd.DataFrame:
    """
    将原始多服务器数据按 (item_name, category, days_open) 聚合：
      avg_price = 所有服务器均价的均值
      min_price = 所有服务器最低价的最小值
    聚合前过滤掉价格无效或挂单数不足 1 件的行。
    """
    df = df.copy()
    df["avg_price"] = pd.to_numeric(df["avg_price"], errors="coerce")
    df["min_price"] = pd.to_numeric(df["min_price"], errors="coerce")
    df["days_open"] = pd.to_numeric(df["days_open"], errors="coerce")
    df["count"]     = pd.to_numeric(df.get("count", pd.Series(dtype=float)),
                                    errors="coerce")

    # 过滤无效行
    df = df[
        df["avg_price"].notna() & (df["avg_price"] > 0) &
        df["min_price"].notna() & (df["min_price"] > 0) &
        df["days_open"].notna() &
        df["count"].notna() & (df["count"] >= 1)
    ].copy()

    market = (
        df.groupby(["item_name", "category", "days_open"], as_index=False)
        .agg(
            avg_price=("avg_price", "mean"),
            min_price=("min_price", "min"),
            n_sources=("avg_price", "count"),
        )
    )
    market["days_open"] = market["days_open"].astype(int)
    return market


def build_analysis(market: pd.DataFrame,
                   df_raw: pd.DataFrame | None = None,
                   fixed_buy_day: int | None = None) -> pd.DataFrame:
    """
    核心计算：对聚合市场快照中每件商品的 (买入天数, 卖出天数) 组合，
    计算利润与收益率。

    fixed_buy_day：若指定，则只以该天数作为买入天数（通常为心动服当前开服天数）；
                   若为 None，则枚举所有可能的买入天数。
    条件：卖出天数 - 买入天数 > 冻结天数
    买入价：买入天数的 min_price（心动服当日最低价）
    卖出价：卖出天数对应的跨服聚合最低价（所有服务器该开服天数最低价的最小值）。
    卖出收入扣除 9% 交易税后计算利润与收益率。
    """
    records = []

    for item, grp in market.groupby("item_name"):
        category    = grp["category"].iloc[0]
        freeze_days = FREEZE_DAYS_BY_CATEGORY.get(category, DEFAULT_FREEZE_DAYS)

        # 按天数排序，建立 {days_open: row} 映射（用于买入价和 days_list）
        grp_sorted = grp.sort_values("days_open")
        day_rows   = {int(r["days_open"]): r for _, r in grp_sorted.iterrows()}
        days_list  = sorted(day_rows.keys())

        # 确定买入天数候选
        if fixed_buy_day is not None:
            buy_days_candidates = [fixed_buy_day] if fixed_buy_day in day_rows else []
        else:
            buy_days_candidates = days_list

        for buy_day in buy_days_candidates:
            buy_row = day_rows[buy_day]
            buy_min = buy_row["min_price"]

            for sell_day in days_list:
                if sell_day - buy_day <= freeze_days:
                    continue

                # 卖出价：该卖出天数的跨服聚合最低价（aggregate_market 已计算）
                sell_price = day_rows[sell_day]["min_price"]
                days_diff  = sell_day - buy_day

                sell_net = sell_price * (1 - TAX_RATE)
                profit   = sell_net - buy_min
                roi      = profit / buy_min * 100

                records.append({
                    "商品名":  item,
                    "分类":    category,
                    "冻结天数": freeze_days,
                    "买入天数": buy_day,
                    "卖出天数": sell_day,
                    "天数差":  days_diff,
                    "买入价":  int(buy_min),
                    "卖出价":  int(sell_price),
                    "利润":    round(profit),
                    "收益率%": round(roi, 1),
                })

    result = pd.DataFrame(records)
    if not result.empty:
        result.sort_values(["商品名", "买入天数", "卖出天数"],
                           inplace=True, ignore_index=True)
    return result


def export_xlsx(tradeable: pd.DataFrame, out_dir: str, date_str: str) -> str:
    """
    从 tradeable_opportunities 中取收益率% 最高的：
      - 5 条冻结 30 天的商品（高级兽诀 / 高级内丹）
      - 5 条冻结  7 天的商品（其余）
    生成美化后的 xlsx 报表。
    """
    top30 = (tradeable[tradeable["冻结天数"] == 30]
             .sort_values("收益率%", ascending=False)
             .head(5))
    top7  = (tradeable[tradeable["冻结天数"] == 7]
             .sort_values("收益率%", ascending=False)
             .head(5))

    # 展示列（中文标题）及其对应的 DataFrame 列名
    DISPLAY = [
        ("商品名",  "商品名"),
        ("分类",    "分类"),
        ("买入天数", "买入天数"),
        ("卖出天数", "卖出天数"),
        ("天数差",  "天数差"),
        ("买入价",  "买入价"),
        ("卖出价(跨服最低价)", "卖出价"),
        ("税后利润", "利润"),
        ("收益率%", "收益率%"),
    ]
    headers   = [d[0] for d in DISPLAY]
    src_cols  = [d[1] for d in DISPLAY]

    # ── 颜色主题 ──────────────────────────────────────────────────────────────
    # 30 天冻结：金色系
    CLR_30_TITLE  = "B8860B"   # 深金
    CLR_30_HEADER = "FFD700"   # 金黄
    CLR_30_ROW_A  = "FFFACD"   # 柠檬绸（浅金）
    CLR_30_ROW_B  = "FFF8DC"   # 玉米丝（更浅）
    # 7 天冻结：蓝色系
    CLR_7_TITLE   = "1B4F8A"   # 深蓝
    CLR_7_HEADER  = "4472C4"   # 标准蓝
    CLR_7_ROW_A   = "DCE6F1"   # 浅蓝
    CLR_7_ROW_B   = "EEF4FB"   # 更浅蓝
    # 正/负利润
    CLR_PROFIT_POS = "1E8449"   # 深绿
    CLR_PROFIT_NEG = "C0392B"   # 深红
    CLR_FONT_DARK  = "1A1A1A"
    CLR_FONT_WHITE = "FFFFFF"

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    def hdr_font(hex_color, bold=True, size=10):
        return Font(name="微软雅黑", bold=bold, color=hex_color, size=size)

    def cell_font(bold=False, size=10, color=CLR_FONT_DARK):
        return Font(name="微软雅黑", bold=bold, color=color, size=size)

    center = Alignment(horizontal="center", vertical="center", wrap_text=False)
    left   = Alignment(horizontal="left",   vertical="center")

    wb = Workbook()
    ws = wb.active
    ws.title = f"套利速览 {date_str}"

    # 固定列宽（9 列：商品名/分类/买入天数/卖出天数/天数差/买入价/卖出价/利润/收益率%）
    COL_WIDTHS = [12, 10, 8, 8, 7, 13, 20, 11, 9]
    for i, w in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    n_cols = len(headers)

    def write_section(start_row: int, df_section: pd.DataFrame,
                      title: str, freeze_days: int) -> int:
        """写一个分组（标题行 + 表头行 + 数据行），返回下一可用行号。"""
        c_title  = CLR_30_TITLE  if freeze_days == 30 else CLR_7_TITLE
        c_header = CLR_30_HEADER if freeze_days == 30 else CLR_7_HEADER
        c_row_a  = CLR_30_ROW_A  if freeze_days == 30 else CLR_7_ROW_A
        c_row_b  = CLR_30_ROW_B  if freeze_days == 30 else CLR_7_ROW_B

        # ── 标题行（合并单元格）──────────────────────────────────────────────
        ws.merge_cells(
            start_row=start_row, start_column=1,
            end_row=start_row,   end_column=n_cols
        )
        title_cell = ws.cell(row=start_row, column=1, value=title)
        title_cell.fill      = fill(c_title)
        title_cell.font      = Font(name="微软雅黑", bold=True,
                                    color=CLR_FONT_WHITE, size=12)
        title_cell.alignment = center
        title_cell.border    = border
        ws.row_dimensions[start_row].height = 22

        # ── 表头行 ────────────────────────────────────────────────────────────
        hdr_row = start_row + 1
        for col_i, h in enumerate(headers, start=1):
            c = ws.cell(row=hdr_row, column=col_i, value=h)
            c.fill      = fill(c_header)
            c.font      = hdr_font(CLR_FONT_WHITE if freeze_days == 7 else CLR_FONT_DARK)
            c.alignment = center
            c.border    = border
        ws.row_dimensions[hdr_row].height = 18

        # ── 数据行 ────────────────────────────────────────────────────────────
        for rank, (_, row) in enumerate(df_section.iterrows(), start=1):
            data_row = hdr_row + rank
            row_fill = fill(c_row_a if rank % 2 == 1 else c_row_b)
            ws.row_dimensions[data_row].height = 17

            for col_i, src in enumerate(src_cols, start=1):
                val = row[src]
                c   = ws.cell(row=data_row, column=col_i, value=val)
                c.fill      = row_fill
                c.border    = border
                c.alignment = left if col_i <= 2 else center

                # 利润列染色
                if src == "利润":
                    if isinstance(val, (int, float)):
                        profit_color = CLR_PROFIT_POS if val >= 0 else CLR_PROFIT_NEG
                        c.font = cell_font(bold=True, color=profit_color)
                    else:
                        c.font = cell_font()
                # 收益率列染色 + 百分号格式
                elif src == "收益率%":
                    if isinstance(val, (int, float)):
                        profit_color = CLR_PROFIT_POS if val >= 0 else CLR_PROFIT_NEG
                        c.font       = cell_font(bold=True, color=profit_color)
                        c.number_format = '0.0"%"'
                    else:
                        c.font = cell_font()
                # 价格列加千位分隔
                elif src in ("买入价", "卖出价"):
                    c.number_format = '#,##0'
                    c.font = cell_font()
                else:
                    c.font = cell_font()

        next_row = hdr_row + len(df_section) + 1
        return next_row

    # ── 写两个分组 ────────────────────────────────────────────────────────────
    row_cursor = 1
    row_cursor = write_section(
        row_cursor, top30,
        f"▶ 高级兽诀 / 高级内丹（冻结 30 天，天数差＞30）",
        freeze_days=30
    )
    row_cursor += 1   # 空一行分隔
    write_section(
        row_cursor, top7,
        f"▶ 其余商品（冻结 7 天）",
        freeze_days=7
    )

    # ── 冻结首列（商品名固定） ────────────────────────────────────────────────
    ws.freeze_panes = "B1"

    out_path = os.path.join(out_dir, "top_opportunities.xlsx")
    wb.save(out_path)
    return out_path
