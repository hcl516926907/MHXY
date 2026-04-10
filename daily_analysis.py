"""
梦幻西游跨服套利分析 —— 每日运行脚本

用法：
  python daily_analysis.py
  启动后输入要分析的日期（默认今天），程序读取该日期及之前所有日期的 data/XXXXXX/ CSV，
  将结果写入 analysis/XXXXXX/。

逻辑：
  聚合所有已采集数据：对每件商品，将所有服务器在相同开服天数下的价格合并：
    avg_price = 所有服务器均价的均值
    min_price = 所有服务器最低价的最小值
  以心动服当前开服天数为买入天数，在所有已采集的卖出天数中寻找套利机会：
    条件：卖出天数 - 买入天数 > 冻结天数
    冻结期规则：
      高级兽诀、高级内丹：30 天
      其余商品：           7 天
  对每对（买入天数，卖出天数），计算：
    min 场景：以买入天数的 min_price 买入，卖出天数的 min_price 卖出
    avg 场景：以买入天数的 avg_price 买入，卖出天数的 avg_price 卖出
  卖出收入扣除 9% 交易税。

输出（保存在 analysis/XXXXXX/）：
  price_analysis.csv          —— 所有套利组合
  tradeable_opportunities.csv —— 同上（所有组合均已满足交易条件），按 min 收益率降序排列
"""

import os
import glob
import re
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import LineChart, Reference, ScatterChart, Series
from openpyxl.chart.layout import Layout, ManualLayout

# ── 配置 ──────────────────────────────────────────────────────────────────────
ROOT_DIR      = os.path.dirname(os.path.abspath(__file__))
DATA_ROOT     = os.path.join(ROOT_DIR, "data")
ANALYSIS_ROOT = os.path.join(ROOT_DIR, "analysis")
HISTORY_XLSX  = os.path.join(ANALYSIS_ROOT, "top_opportunities_history.xlsx")

TAX_RATE = 0.09

# 冻结天数（按商品分类）；不在表中的默认 7 天
FREEZE_DAYS_BY_CATEGORY: dict[str, int] = {
    "高级兽诀": 30,
    "高级内丹": 30,
}
DEFAULT_FREEZE_DAYS = 7
# ─────────────────────────────────────────────────────────────────────────────


def prompt_date() -> str:
    """提示用户输入日期，默认为今天。返回 YYYYMMDD 格式字符串。"""
    today = datetime.now().strftime("%Y%m%d")
    while True:
        raw = input(f"请输入分析日期（格式 YYYYMMDD，直接回车默认 {today}）：").strip()
        if raw == "":
            return today
        if re.fullmatch(r"\d{8}", raw):
            return raw
        print("  格式错误，请输入 8 位数字，例如 20260403。")


def load_all_source_csvs(up_to_date: str) -> pd.DataFrame:
    """
    读取 DATA_ROOT 下所有日期目录（≤ up_to_date）中的全部 CSV 文件。
    up_to_date 为 YYYYMMDD 格式字符串。
    """
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

    print(f"  已读取 {len(date_dirs)} 个日期目录（{date_dirs[0]} ~ {date_dirs[-1]}）")
    return pd.concat(frames, ignore_index=True)


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
                   fixed_buy_day: int | None = None) -> pd.DataFrame:
    """
    核心计算：对聚合市场快照中每件商品的 (买入天数, 卖出天数) 组合，
    计算 min/avg 两个场景的利润与收益率。

    fixed_buy_day：若指定，则只以该天数作为买入天数（通常为心动服当前开服天数）；
                   若为 None，则枚举所有可能的买入天数。
    条件：卖出天数 - 买入天数 > 冻结天数
    min 场景：买入天数的 min_price 买入，卖出天数的 min_price 卖出
    avg 场景：买入天数的 avg_price 买入，卖出天数的 avg_price 卖出
    """
    records = []

    for item, grp in market.groupby("item_name"):
        category    = grp["category"].iloc[0]
        freeze_days = FREEZE_DAYS_BY_CATEGORY.get(category, DEFAULT_FREEZE_DAYS)

        # 按天数排序，建立 {days_open: row} 映射
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
            buy_avg = buy_row["avg_price"]

            for sell_day in days_list:
                if sell_day - buy_day <= freeze_days:
                    continue

                sell_row = day_rows[sell_day]
                sell_min = sell_row["min_price"]
                sell_avg = sell_row["avg_price"]

                days_diff = sell_day - buy_day

                # min 场景
                sell_net_min = sell_min * (1 - TAX_RATE)
                profit_min   = sell_net_min - buy_min
                roi_min      = profit_min / buy_min * 100

                # avg 场景
                sell_net_avg = sell_avg * (1 - TAX_RATE)
                profit_avg   = sell_net_avg - buy_avg
                roi_avg      = profit_avg / buy_avg * 100

                records.append({
                    "商品名":       item,
                    "分类":         category,
                    "冻结天数":     freeze_days,
                    "买入天数":     buy_day,
                    "卖出天数":     sell_day,
                    "天数差":       days_diff,
                    "买入价(最低)": int(buy_min),
                    "买入价(均价)": int(buy_avg),
                    "卖出价(最低)": int(sell_min),
                    "卖出价(均价)": int(sell_avg),
                    "税后收入_min": round(sell_net_min),
                    "税后收入_avg": round(sell_net_avg),
                    "利润_min":     round(profit_min),
                    "利润_avg":     round(profit_avg),
                    "收益率%_min":  round(roi_min, 1),
                    "收益率%_avg":  round(roi_avg, 1),
                })

    result = pd.DataFrame(records)
    if not result.empty:
        result.sort_values(["商品名", "买入天数", "卖出天数"],
                           inplace=True, ignore_index=True)
    return result


def export_xlsx(tradeable: pd.DataFrame, out_dir: str, date_str: str) -> str:
    """
    从 tradeable_opportunities 中取收益率%_min 最高的：
      - 5 条冻结 30 天的商品（高级兽诀 / 高级内丹）
      - 5 条冻结  7 天的商品（其余）
    生成美化后的 xlsx 报表。
    """
    top30 = (tradeable[tradeable["冻结天数"] == 30]
             .sort_values("收益率%_min", ascending=False)
             .head(5))
    top7  = (tradeable[tradeable["冻结天数"] == 7]
             .sort_values("收益率%_min", ascending=False)
             .head(5))

    # 展示列（中文标题）及其对应的 DataFrame 列名
    DISPLAY = [
        ("商品名",       "商品名"),
        ("分类",         "分类"),
        ("买入天数",     "买入天数"),
        ("卖出天数",     "卖出天数"),
        ("天数差",       "天数差"),
        ("买入价(最低)", "买入价(最低)"),
        ("卖出价(最低)", "卖出价(最低)"),
        ("税后利润",     "利润_min"),
        ("收益率%",      "收益率%_min"),
        ("买入价(均价)", "买入价(均价)"),
        ("卖出价(均价)", "卖出价(均价)"),
        ("税后利润(均)", "利润_avg"),
        ("收益率%(均)",  "收益率%_avg"),
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

    # 固定列宽（13 列：商品名/分类/买入天数/卖出天数/天数差/买入min/卖出min/利润/ROI/买入avg/卖出avg/利润均/ROI均）
    COL_WIDTHS = [12, 10, 8, 8, 7, 13, 13, 11, 9, 13, 13, 11, 9]
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

        # 分隔线：最低价区 / 均价区
        # 列 1-9 = 最低价组，10-13 = 均价组
        # 在表头加视觉分隔（左边框加粗）
        sep_col = 10  # 均价区起始列
        thick = Side(style="medium", color="888888")
        for r in range(start_row, start_row + 2 + len(df_section)):
            cell = ws.cell(row=r, column=sep_col)
            cell.border = Border(
                left=thick,
                right=thin, top=thin, bottom=thin
            )

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
                if src in ("利润_min", "利润_avg"):
                    if isinstance(val, (int, float)):
                        profit_color = CLR_PROFIT_POS if val >= 0 else CLR_PROFIT_NEG
                        c.font = cell_font(bold=True, color=profit_color)
                    else:
                        c.font = cell_font()
                # 收益率列染色 + 百分号格式
                elif src in ("收益率%_min", "收益率%_avg"):
                    if isinstance(val, (int, float)):
                        profit_color = CLR_PROFIT_POS if val >= 0 else CLR_PROFIT_NEG
                        c.font       = cell_font(bold=True, color=profit_color)
                        c.number_format = '0.0"%"'
                    else:
                        c.font = cell_font()
                # 价格列加千位分隔
                elif src in ("买入价(最低)", "卖出价(最低)", "税后收入_min",
                             "买入价(均价)", "卖出价(均价)", "税后收入_avg"):
                    c.number_format = '#,##0'
                    c.font = cell_font()
                else:
                    c.font = cell_font()

            # 分隔线延伸到数据行
            ws.cell(row=data_row, column=sep_col).border = Border(
                left=thick, right=thin, top=thin, bottom=thin
            )

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


# ── 历史汇总相关常量 ──────────────────────────────────────────────────────────
# 历史数据 sheet 列头（按顺序）
_HIST_HEADERS = [
    "日期", "冻结天数", "排名",
    "商品名", "分类", "买入天数", "卖出天数", "天数差",
    "买入价(最低)", "卖出价(最低)", "税后利润", "收益率%",
    "买入价(均价)", "卖出价(均价)", "税后利润(均)", "收益率%(均)",
]
# tradeable_opportunities.csv 中对应列（前 3 个由代码填充）
_HIST_SRC = [
    None, None, None,
    "商品名", "分类", "买入天数", "卖出天数", "天数差",
    "买入价(最低)", "卖出价(最低)", "利润_min", "收益率%_min",
    "买入价(均价)", "卖出价(均价)", "利润_avg", "收益率%_avg",
]
# Sheet 1 展示列（去掉日期/冻结天数/排名），及对应 历史数据 sheet 列字母
_DISP_HEADERS = _HIST_HEADERS[3:]   # 13 列
_HIST_DATA_COLS = [get_column_letter(i) for i in range(4, 4 + len(_DISP_HEADERS))]  # D~P


def _collect_all_tops() -> list:
    """
    扫描 analysis/XXXXXX/ 下所有 tradeable_opportunities.csv，
    提取每日 top5×2（30天/7天冻结），加上日期、冻结天数、排名列。
    返回按 日期→冻结天数→排名 排序的字典列表。
    """
    all_rows = []
    if not os.path.isdir(ANALYSIS_ROOT):
        return all_rows
    for entry in sorted(os.listdir(ANALYSIS_ROOT)):
        if not re.fullmatch(r"\d{8}", entry):
            continue
        csv_path = os.path.join(ANALYSIS_ROOT, entry, "tradeable_opportunities.csv")
        if not os.path.isfile(csv_path):
            continue
        try:
            df = pd.read_csv(csv_path, encoding="utf-8-sig")
        except Exception as e:
            print(f"  [警告] 跳过 {entry}: {e}")
            continue

        date_int = int(entry)
        for freeze in [30, 7]:
            subset = (df[df["冻结天数"] == freeze]
                      .sort_values("收益率%_min", ascending=False)
                      .head(5))
            for rank, (_, row) in enumerate(subset.iterrows(), start=1):
                record = {"日期": date_int, "冻结天数": freeze, "排名": rank}
                for src_col, hist_hdr in zip(_HIST_SRC[3:], _HIST_HEADERS[3:]):
                    val = row.get(src_col, "")
                    record[hist_hdr] = "" if pd.isna(val) else val
                all_rows.append(record)
    return all_rows


def build_history_xlsx():
    """
    生成 analysis/top_opportunities_history.xlsx，含两个 sheet：
      - 套利速览：日期下拉选择 + SUMPRODUCT 公式动态展示当日 top5×2
      - 历史数据：所有历史记录原始行
    """
    rows = _collect_all_tops()
    if not rows:
        print("  [警告] 未找到任何历史数据，跳过历史汇总。")
        return

    all_dates = sorted({r["日期"] for r in rows})
    latest_date = all_dates[-1]

    # ── 配色（与 export_xlsx 保持一致） ───────────────────────────────────────
    CLR_30_TITLE   = "B8860B"; CLR_30_HEADER  = "FFD700"
    CLR_30_ROW_A   = "FFFACD"; CLR_30_ROW_B   = "FFF8DC"
    CLR_7_TITLE    = "1B4F8A"; CLR_7_HEADER   = "4472C4"
    CLR_7_ROW_A    = "DCE6F1"; CLR_7_ROW_B    = "EEF4FB"
    CLR_PROFIT_POS = "1E8449"; CLR_PROFIT_NEG = "C0392B"
    CLR_FONT_DARK  = "1A1A1A"; CLR_FONT_WHITE = "FFFFFF"

    thin   = Side(style="thin",   color="BFBFBF")
    thick  = Side(style="medium", color="888888")
    b_all  = Border(left=thin,  right=thin,  top=thin, bottom=thin)
    b_sep  = Border(left=thick, right=thin,  top=thin, bottom=thin)

    def fill(hex_): return PatternFill("solid", fgColor=hex_)
    def myfont(bold=False, size=10, color=CLR_FONT_DARK):
        return Font(name="微软雅黑", bold=bold, color=color, size=size)

    center = Alignment(horizontal="center", vertical="center")
    left_  = Alignment(horizontal="left",   vertical="center")

    wb = Workbook()

    # ════════════════════════════════════════════════════════════════════════
    # Sheet 1: 套利速览
    # ════════════════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "套利速览"

    n_disp = len(_DISP_HEADERS)   # 13
    COL_WIDTHS = [12, 10, 8, 8, 7, 13, 13, 11, 9, 13, 13, 11, 9]
    for i, w in enumerate(COL_WIDTHS, start=1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    # ── 日期选择行 ────────────────────────────────────────────────────────
    label = ws1["A1"]
    label.value = "选择日期："
    label.fill  = fill("E8F0FE")
    label.font  = myfont(bold=True, size=11)
    label.alignment = Alignment(horizontal="right", vertical="center")
    label.border = b_all

    ws1.merge_cells("B1:C1")
    date_cell = ws1["B1"]
    date_cell.value     = latest_date
    date_cell.fill      = fill("FFFFFF")
    date_cell.font      = myfont(bold=True, size=12, color="1B4F8A")
    date_cell.alignment = center
    date_cell.border    = b_all
    ws1.row_dimensions[1].height = 26

    # 下拉数据验证
    date_list_str = ",".join(str(d) for d in all_dates)
    dv = DataValidation(type="list", formula1=f'"{date_list_str}"', allow_blank=True)
    dv.sqref = "B1"
    ws1.add_data_validation(dv)

    sep_col = 10   # 均价区起始（1-based）

    def _sumproduct_formula(freeze: int, rank: int, hist_col: str) -> str:
        """生成 SUMPRODUCT+INDEX 公式，从 历史数据 sheet 中查询对应单元格值。"""
        rng = "历史数据!$A$2:$A$5000"
        cond = (
            f"(历史数据!$A$2:$A$5000=$B$1)"
            f"*(历史数据!$B$2:$B$5000={freeze})"
            f"*(历史数据!$C$2:$C$5000={rank})"
        )
        return (
            f'=IFERROR(IF(SUMPRODUCT({cond})=0,"",'
            f'INDEX(历史数据!${hist_col}:${hist_col},'
            f'SUMPRODUCT({cond}*ROW(历史数据!$A$2:$A$5000)))),"")'
        )

    def write_section_formulas(start_row: int, freeze: int, title: str) -> int:
        c_title  = CLR_30_TITLE  if freeze == 30 else CLR_7_TITLE
        c_header = CLR_30_HEADER if freeze == 30 else CLR_7_HEADER
        c_row_a  = CLR_30_ROW_A  if freeze == 30 else CLR_7_ROW_A
        c_row_b  = CLR_30_ROW_B  if freeze == 30 else CLR_7_ROW_B

        # 标题行
        ws1.merge_cells(start_row=start_row, start_column=1,
                        end_row=start_row, end_column=n_disp)
        tc = ws1.cell(row=start_row, column=1, value=title)
        tc.fill = fill(c_title)
        tc.font = Font(name="微软雅黑", bold=True, color=CLR_FONT_WHITE, size=12)
        tc.alignment = center
        tc.border = b_all
        ws1.row_dimensions[start_row].height = 22

        # 表头行
        hdr_row = start_row + 1
        for ci, h in enumerate(_DISP_HEADERS, start=1):
            c = ws1.cell(row=hdr_row, column=ci, value=h)
            c.fill = fill(c_header)
            c.font = Font(name="微软雅黑", bold=True,
                          color=CLR_FONT_WHITE if freeze == 7 else CLR_FONT_DARK,
                          size=10)
            c.alignment = center
            c.border = b_sep if ci == sep_col else b_all
        ws1.row_dimensions[hdr_row].height = 18

        # 数据行（公式）
        for rank in range(1, 6):
            dr = hdr_row + rank
            rf = fill(c_row_a if rank % 2 == 1 else c_row_b)
            ws1.row_dimensions[dr].height = 17
            for ci, (hist_col, disp_hdr) in enumerate(
                    zip(_HIST_DATA_COLS, _DISP_HEADERS), start=1):
                formula = _sumproduct_formula(freeze, rank, hist_col)
                c = ws1.cell(row=dr, column=ci, value=formula)
                c.fill      = rf
                c.border    = b_sep if ci == sep_col else b_all
                c.alignment = left_ if ci <= 2 else center
                if disp_hdr in ("税后利润", "税后利润(均)"):
                    c.number_format = '#,##0'
                    c.font = myfont()
                elif disp_hdr in ("收益率%", "收益率%(均)"):
                    c.number_format = '0.0"%"'
                    c.font = myfont()
                elif disp_hdr in ("买入价(最低)", "卖出价(最低)", "买入价(均价)", "卖出价(均价)"):
                    c.number_format = '#,##0'
                    c.font = myfont()
                else:
                    c.font = myfont()

        return hdr_row + 6  # 下一个可用行

    row_cursor = 3   # 第 2 行空出，从第 3 行开始写第一个分组
    row_cursor = write_section_formulas(
        row_cursor, 30, "▶ 高级兽诀 / 高级内丹（冻结 30 天）")
    row_cursor += 1  # 空一行
    write_section_formulas(
        row_cursor, 7,  "▶ 其余商品（冻结 7 天）")

    ws1.freeze_panes = "A2"

    # ════════════════════════════════════════════════════════════════════════
    # Sheet 2: 历史数据
    # ════════════════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("历史数据")

    HDR_BG = "2E4057"
    HIST_COL_WIDTHS = [10, 8, 5, 12, 10, 8, 8, 7, 13, 13, 11, 9, 13, 13, 11, 9]
    for i, w in enumerate(HIST_COL_WIDTHS, start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    for ci, h in enumerate(_HIST_HEADERS, start=1):
        c = ws2.cell(row=1, column=ci, value=h)
        c.fill = fill(HDR_BG)
        c.font = Font(name="微软雅黑", bold=True, color=CLR_FONT_WHITE, size=10)
        c.alignment = center
        c.border = b_all
    ws2.row_dimensions[1].height = 20

    stripe_a = fill("F7F9FC")
    stripe_b = fill("FFFFFF")
    for ri, row in enumerate(rows, start=2):
        rf = stripe_a if ri % 2 == 0 else stripe_b
        ws2.row_dimensions[ri].height = 16
        for ci, h in enumerate(_HIST_HEADERS, start=1):
            val = row.get(h, "")
            c = ws2.cell(row=ri, column=ci, value=val)
            c.fill      = rf
            c.border    = b_all
            c.alignment = left_ if h == "商品名" else center
            if h in ("税后利润", "税后利润(均)"):
                if isinstance(val, (int, float)):
                    c.font = Font(name="微软雅黑", bold=True, size=10,
                                  color=CLR_PROFIT_POS if val >= 0 else CLR_PROFIT_NEG)
                else:
                    c.font = myfont()
                c.number_format = '#,##0'
            elif h in ("收益率%", "收益率%(均)"):
                if isinstance(val, (int, float)):
                    c.font = Font(name="微软雅黑", bold=True, size=10,
                                  color=CLR_PROFIT_POS if val >= 0 else CLR_PROFIT_NEG)
                else:
                    c.font = myfont()
                c.number_format = '0.0"%"'
            elif h in ("买入价(最低)", "卖出价(最低)", "买入价(均价)", "卖出价(均价)"):
                c.number_format = '#,##0'
                c.font = myfont()
            else:
                c.font = myfont()

    ws2.freeze_panes = "A2"

    os.makedirs(ANALYSIS_ROOT, exist_ok=True)
    wb.save(HISTORY_XLSX)
    print(f"历史汇总已保存：analysis/top_opportunities_history.xlsx"
          f"（{len(all_dates)} 天，共 {len(rows)} 行）")


TREND_XLSX = os.path.join(ANALYSIS_ROOT, "price_trend.xlsx")  # 基础路径（无 VBA 时使用）

# 折线图配色（最多支持 16 个服务器，用于未知服务器的 fallback）
_SERIES_COLORS = [
    "4472C4", "ED7D31", "70AD47", "FF0000", "FFC000",
    "5B9BD5", "A9D18E", "264478", "9E480E", "C55A11",
    "843C0C", "7030A0", "00B0F0", "92D050", "FF66FF", "595959",
]

# 固定服务器配色（折线 + 标记边框统一使用）
_SERVER_COLORS = {
    "大吉大利": "0072B2",
    "天命":     "E69F00",
    "心动":     "009E73",
    "飞天":     "CC3311",
}

def _get_server_color(name: str) -> str:
    """返回服务器的固定颜色；未知服务器 fallback 到 _SERIES_COLORS。"""
    return _SERVER_COLORS.get(name, _SERIES_COLORS[abs(hash(name)) % len(_SERIES_COLORS)])


_VBA_CHART_AXIS = """\
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("B1:B2")) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    On Error GoTo done
    Application.Calculate
    UpdateYAxis
done:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Activate()
    On Error Resume Next
    Application.Calculate
    UpdateYAxis
End Sub

Private Sub UpdateYAxis()
    ' 动态定位数据区域：先找含"开服天数"的表头行，再从下一行起扫描
    If Me.ChartObjects.Count = 0 Then Exit Sub

    ' 找表头行（A列中值为"开服天数"的行）
    Dim hdrRow As Long
    hdrRow = 0
    Dim i As Long
    For i = 1 To 200
        If CStr(Me.Cells(i, 1).Value) = "开服天数" Then
            hdrRow = i
            Exit For
        End If
    Next i
    If hdrRow = 0 Then Exit Sub

    Dim lastRow As Long, lastCol As Long
    lastRow = Me.Cells(Me.Rows.Count, 1).End(xlUp).Row   ' 列A最后一行 = 最后一个开服天数
    lastCol = Me.Cells(hdrRow, Me.Columns.Count).End(xlToLeft).Column  ' 表头行最后一列

    Dim minVal As Double, maxVal As Double, hasData As Boolean
    Dim r As Long, c As Long, v As Variant
    minVal = 1E+15
    maxVal = -1E+15
    hasData = False

    For r = hdrRow + 1 To lastRow
        For c = 2 To lastCol
            v = Me.Cells(r, c).Value
            If IsNumeric(v) And Not IsEmpty(v) Then
                If CDbl(v) < minVal Then minVal = CDbl(v)
                If CDbl(v) > maxVal Then maxVal = CDbl(v)
                hasData = True
            End If
        Next c
    Next r

    If Not hasData Then Exit Sub
    If minVal >= maxVal Then Exit Sub

    On Error Resume Next
    With Me.ChartObjects(1).Chart.Axes(2)
        .MinimumScale = minVal * 0.95
        .MaximumScale = maxVal * 1.05
    End With
End Sub
"""


def _try_inject_vba(src_xlsx: str, out_xlsm: str) -> bool:
    """用 Excel COM 向"图表数据"工作表注入 VBA，另存为 .xlsm。
    src_xlsx 是已存在的有效 xlsx 文件，注入成功后另存为 out_xlsm。
    返回 True 表示成功；返回 False 表示失败（src_xlsx 保持不变）。
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print("  [警告] win32com 不可用，Y 轴将回退为静态范围（.xlsx）。")
        return False

    pythoncom.CoInitialize()
    xl = None
    wbk = None
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False

        wbk = xl.Workbooks.Open(os.path.abspath(src_xlsx))
        ws = wbk.Worksheets("图表数据")

        try:
            vb_comp = wbk.VBProject.VBComponents(ws.CodeName)
        except Exception:
            raise RuntimeError(
                "无法访问 VBProject\n"
                "  → 请在 Excel → 选项 → 信任中心 → 宏设置 中\n"
                "    勾选「信任对 VBA 工程对象模型的访问」后重新生成。"
            )

        mod = vb_comp.CodeModule
        if mod.CountOfLines > 0:
            mod.DeleteLines(1, mod.CountOfLines)
        mod.AddFromString(_VBA_CHART_AXIS)

        # 若目标 xlsm 已存在，先删除，避免 SaveAs 覆盖已有宏工作簿时报错
        if os.path.exists(out_xlsm):
            os.remove(out_xlsm)
        # FileFormat 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
        wbk.SaveAs(os.path.abspath(out_xlsm), FileFormat=52)
        print("  VBA 注入成功：Y 轴将随商品选择动态缩放（.xlsm）。")
        return True
    except Exception as e:
        print(f"  [警告] VBA 注入失败：{e}")
        print("  → 已回退为静态 Y 轴范围，文件保存为 .xlsx。")
        return False
    finally:
        if wbk is not None:
            try:
                wbk.Close(False)
            except Exception:
                pass
        if xl is not None:
            try:
                xl.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _post_process_chart_xml(xlsx_path: str, n_servers: int) -> None:
    """
    直接修改 openpyxl 保存的 xlsx 内 chart XML，添加：
    - 图表区 + 绘图区白色背景
    - y 轴主网格线：浅灰（#D9D9D9）虚线
    - 两轴轴线颜色：中灰（#808080）
    """
    import zipfile, io, re as _re

    with open(xlsx_path, "rb") as f:
        buf = io.BytesIO(f.read())

    with zipfile.ZipFile(buf, "r") as zin:
        names = zin.namelist()
        files = {n: zin.read(n) for n in names}

    chart_files = sorted(
        n for n in names if _re.match(r"xl/charts/chart\d+\.xml", n)
    )
    if not chart_files:
        return

    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    SOLID_WHITE = f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="FFFFFF"/></a:solidFill>'
    NO_LINE     = f'<a:ln xmlns:a="{A}"><a:noFill/></a:ln>'
    GRAY_LINE   = (f'<a:ln xmlns:a="{A}">'
                   f'<a:solidFill><a:srgbClr val="808080"/></a:solidFill>'
                   f'</a:ln>')

    xml = files[chart_files[0]].decode("utf-8")

    # 检测 chart 命名空间前缀（openpyxl 通常输出 c:，也可能为空前缀）
    c_ns_match = _re.search(
        r'xmlns:([A-Za-z0-9_]+)="http://schemas\.openxmlformats\.org/drawingml/2006/chart"',
        xml
    )
    c = (c_ns_match.group(1) + ":") if c_ns_match else ""

    GRID_SPPR = (f'<{c}spPr>'
                 f'<a:ln xmlns:a="{A}">'
                 f'<a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill>'
                 f'<a:prstDash val="sysDash"/>'
                 f'</a:ln></{c}spPr>')

    # 1. 图表区背景（chartSpace 第一个子元素之前插入 spPr）
    xml = xml.replace(f"<{c}chart>",
                      f"<{c}spPr>{SOLID_WHITE}{NO_LINE}</{c}spPr><{c}chart>", 1)

    # 2. 绘图区背景（plotArea 开头）
    xml = xml.replace(f"<{c}plotArea>",
                      f"<{c}plotArea><{c}spPr>{SOLID_WHITE}</{c}spPr>", 1)

    # 3. y 轴主网格线样式（替换第一个 <majorGridlines/> 为带样式版本）
    xml = xml.replace(f"<{c}majorGridlines/>",
                      f"<{c}majorGridlines>{GRID_SPPR}</{c}majorGridlines>", 1)

    # 4. 轴线颜色：在每个 </valAx> 之前插入 spPr（若尚无 spPr）
    xml = _re.sub(
        rf"(<{c}crossAx[^/]*/>\s*)(</{c}valAx>)",
        r"\1" + f"<{c}spPr>{GRAY_LINE}</{c}spPr>" + r"\2",
        xml
    )

    files[chart_files[0]] = xml.encode("utf-8")

    buf2 = io.BytesIO()
    with zipfile.ZipFile(buf2, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    buf2.seek(0)
    with open(xlsx_path, "wb") as f:
        f.write(buf2.read())


def build_price_trend_xlsx(df_raw: pd.DataFrame) -> None:
    """
    生成 analysis/price_trend.xlsm，包含：
      - 图表数据 sheet（第一个）：选择商品/价格类型后，用公式展示各服务器价格矩阵
        并嵌入折线图（x=开服天数，不同颜色=不同服务器，不同形状标记=不同星期）
      - 数据 sheet（第二个）：未聚合的原始价格记录（每服务器/开服天数一行）
      - 商品列表 sheet（隐藏）：用于下拉验证
    """
    # ── 数据过滤 ──────────────────────────────────────────────────────────────
    df = df_raw.copy()
    for col in ("avg_price", "min_price", "days_open", "count"):
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df[
        df["avg_price"].notna() & (df["avg_price"] > 0) &
        df["min_price"].notna() & (df["min_price"] > 0) &
        df["days_open"].notna() &
        df["count"].notna() & (df["count"] >= 1)
    ].copy()
    df["days_open"] = df["days_open"].astype(int)
    df["avg_price"] = df["avg_price"].round().astype(int)
    df["min_price"] = df["min_price"].round().astype(int)

    # 从 timestamp 计算星期（0=周一，6=周日）
    if "timestamp" in df.columns:
        df["_ts_parsed"] = pd.to_datetime(df["timestamp"], errors="coerce", format="mixed")
        df["dow"] = df["_ts_parsed"].dt.dayofweek
        df["dow"] = df["dow"].fillna(-1).astype(int)
    else:
        df["_ts_parsed"] = pd.NaT
        df["dow"] = -1

    # 同一(商品,服务器,开服天数)保留最新一条
    if "timestamp" in df.columns:
        df = df.sort_values("_ts_parsed")
    df = df.drop_duplicates(
        subset=["item_name", "server_name", "days_open"], keep="last"
    )
    df = df.drop(columns=["_ts_parsed"])

    # 按 (父类拼音, 商品名拼音) 升序排列商品
    item_category = (
        df.groupby("item_name")["category"]
        .agg(lambda x: x.mode().iloc[0])
        .to_dict()
    )
    try:
        from pypinyin import lazy_pinyin
        def _pinyin_key(text: str):
            return lazy_pinyin(text)
        items = sorted(
            df["item_name"].unique(),
            key=lambda x: (_pinyin_key(item_category.get(x, "")), _pinyin_key(x))
        )
    except ImportError:
        items = sorted(
            df["item_name"].unique(),
            key=lambda x: (item_category.get(x, ""), x)
        )

    servers = sorted(df["server_name"].unique())
    days    = sorted(df["days_open"].unique())

    if not items:
        print("  [警告] 无有效数据，跳过 price_trend.xlsx 生成。")
        return

    n_items   = len(items)
    n_servers = len(servers)
    n_days    = len(days)
    data_end  = len(df) + 1   # 数据 sheet 最后数据行（含表头偏移）

    # ── 样式工具 ──────────────────────────────────────────────────────────────
    thin   = Side(style="thin",   color="BFBFBF")
    b_all  = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    left_a = Alignment(horizontal="left",   vertical="center")

    def fill(hex_): return PatternFill("solid", fgColor=hex_)
    def myfont(bold=False, size=10, color="1A1A1A"):
        return Font(name="微软雅黑", bold=bold, color=color, size=size)

    # 星期相关常量（圆/方/三角/菱形/×/加/五角星 → 周一~周日）
    DOW_LABELS   = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
    DOW_MARKERS  = ["circle", "square", "triangle", "diamond", "x", "plus", "star"]
    # 描边标记（只有边线，无实心区域）：x, plus, star
    DOW_STROKED  = {}
    # 标记辅助列起始（1-based，紧接服务器列之后）
    MARKER_START_COL = n_servers + 2

    wb = Workbook()

    # ════════════════════════════════════════════════════════════════════════
    # Sheet 1: 图表数据（选择商品/价格类型 → 公式矩阵 + 折线图）
    # ════════════════════════════════════════════════════════════════════════
    ws_chart = wb.active
    ws_chart.title = "图表数据"

    # ── 控制区（行 1-2）────────────────────────────────────────────────────
    CTRL_BG = "E8F0FE"
    for label, row_i, default, dv_formula in [
        ("商品名：",   1, items[0],  f"=商品列表!$A$1:$A${n_items}"),
        ("价格类型：", 2, "最低价",  '"最低价,均价"'),
    ]:
        lc = ws_chart.cell(row=row_i, column=1, value=label)
        lc.fill = fill(CTRL_BG)
        lc.font = myfont(bold=True, size=11)
        lc.alignment = Alignment(horizontal="right", vertical="center")
        lc.border = b_all

        vc = ws_chart.cell(row=row_i, column=2, value=default)
        vc.fill = fill("FFFFFF")
        vc.font = myfont(bold=True, size=11, color="1B4F8A")
        vc.alignment = center
        vc.border = b_all
        ws_chart.row_dimensions[row_i].height = 22

        dv = DataValidation(type="list", formula1=dv_formula, allow_blank=False)
        dv.sqref = f"B{row_i}"
        ws_chart.add_data_validation(dv)

    ws_chart.column_dimensions["A"].width = 14
    ws_chart.column_dimensions["B"].width = 18

    # ── 矩阵表头（图表下方）──────────────────────────────────────────────────
    # 图表锚定在 A4，高度 14 cm ≈ 27 行（默认 15pt 行高 0.529 cm/行），留 5 行间距
    HDR_ROW = 36
    hc = ws_chart.cell(row=HDR_ROW, column=1, value="开服天数")
    hc.fill = fill("2E4057")
    hc.font = myfont(bold=True, color="FFFFFF")
    hc.alignment = center
    hc.border = b_all
    ws_chart.row_dimensions[HDR_ROW].height = 20

    for ci, srv in enumerate(servers, start=2):
        c = ws_chart.cell(row=HDR_ROW, column=ci, value=srv)
        c.fill = fill(_get_server_color(srv))
        c.font = myfont(bold=True, color="FFFFFF")
        c.alignment = center
        c.border = b_all
        ws_chart.column_dimensions[get_column_letter(ci)].width = 12

    # 标记辅助列表头（服务器 × 星期，列宽压窄）
    for si, srv in enumerate(servers):
        for dj, dow_label in enumerate(DOW_LABELS):
            mc = MARKER_START_COL + si * 7 + dj
            c = ws_chart.cell(row=HDR_ROW, column=mc, value=dow_label[1])
            c.fill = fill(_get_server_color(srv))
            c.font = myfont(bold=False, size=8, color="FFFFFF")
            c.alignment = center
            c.border = b_all
            ws_chart.column_dimensions[get_column_letter(mc)].width = 5

    # ── 矩阵数据（行 5 开始，每行一个开服天数）─────────────────────────────
    FIRST_DATA_ROW = HDR_ROW + 1
    for ri_off, day in enumerate(days):
        ri = FIRST_DATA_ROW + ri_off
        # 开服天数列
        dc = ws_chart.cell(row=ri, column=1, value=day)
        dc.fill = fill("F2F2F2")
        dc.font = myfont(bold=True)
        dc.alignment = center
        dc.border = b_all
        ws_chart.row_dimensions[ri].height = 16

        # 服务器数据列
        for ci, srv in enumerate(servers, start=2):
            srv_col = get_column_letter(ci)
            # 若该服务器/天数无数据则返回 NA()（图表中显示为断点）
            formula = (
                f"=IFERROR(IF("
                f"SUMPRODUCT((数据!$A$2:$A${data_end}=$B$1)"
                f"*(数据!$B$2:$B${data_end}={srv_col}${HDR_ROW})"
                f"*(数据!$C$2:$C${data_end}=$A{ri}))=0,"
                f"NA(),"
                f"SUMPRODUCT((数据!$A$2:$A${data_end}=$B$1)"
                f"*(数据!$B$2:$B${data_end}={srv_col}${HDR_ROW})"
                f"*(数据!$C$2:$C${data_end}=$A{ri})"
                f"*IF($B$2=\"最低价\",数据!$D$2:$D${data_end},数据!$E$2:$E${data_end}))"
                f"),NA())"
            )
            c = ws_chart.cell(row=ri, column=ci, value=formula)
            c.border = b_all
            c.alignment = center
            c.font = myfont()
            c.number_format = "#,##0"

        # 标记辅助列（每服务器 × 每星期）：返回对应（商品,服务器,开服天数,星期）的价格
        for si, srv in enumerate(servers):
            srv_col_letter = get_column_letter(si + 2)
            for dj in range(7):
                mc = MARKER_START_COL + si * 7 + dj
                formula = (
                    f"=IFERROR(IF("
                    f"SUMPRODUCT((数据!$A$2:$A${data_end}=$B$1)"
                    f"*(数据!$B$2:$B${data_end}={srv_col_letter}${HDR_ROW})"
                    f"*(数据!$C$2:$C${data_end}=$A{ri})"
                    f"*(数据!$F$2:$F${data_end}={dj}))=0,"
                    f"NA(),"
                    f"SUMPRODUCT((数据!$A$2:$A${data_end}=$B$1)"
                    f"*(数据!$B$2:$B${data_end}={srv_col_letter}${HDR_ROW})"
                    f"*(数据!$C$2:$C${data_end}=$A{ri})"
                    f"*(数据!$F$2:$F${data_end}={dj})"
                    f"*IF($B$2=\"最低价\",数据!$D$2:$D${data_end},数据!$E$2:$E${data_end}))"
                    f"),NA())"
                )
                c = ws_chart.cell(row=ri, column=mc, value=formula)
                c.border = b_all
                c.alignment = center
                c.font = myfont(size=8)
                c.number_format = "#,##0"

    # ── 折线图（ScatterChart，x 轴为数值轴，正确显示开服天数）───────────────
    chart = ScatterChart()
    chart.title       = "商品价格随开服天数变化（请在 B1/B2 选择商品和价格类型）"
    chart.scatterStyle = "lineMarker"
    chart.style       = None          # 不套用内置样式
    chart.y_axis.title = "价格"
    chart.x_axis.title = "开服天数"
    chart.y_axis.numFmt = "#,##0"
    chart.x_axis.numFmt = "0"
    chart.height = 14
    chart.width  = max(24, n_servers * 3 + 18)

    # ── x 轴范围：固定为全局开服天数范围（与商品无关）──────────────────────────
    chart.x_axis.scaling.min = float(min(days))
    chart.x_axis.scaling.max = float(max(days))
    # y 轴初始范围：取默认商品（items[0]）最低价数据；VBA 会在用户切换商品后动态更新
    _init_prices = (
        df[df["item_name"] == items[0]]["min_price"].dropna().tolist()
    )
    _init_prices = [p for p in _init_prices if isinstance(p, (int, float)) and p > 0]
    if _init_prices:
        chart.y_axis.scaling.min = float(min(_init_prices)) * 0.95
        chart.y_axis.scaling.max = float(max(_init_prices)) * 1.05

    # 轴线锚定在数据边缘，避免默认 crosses=autoZero 时轴线跑到 y=0（超出可见区域）
    # 导致 tickLblPos="low" 把刻度标签渲染到图表底部最外侧
    chart.x_axis.crosses = "min"   # x 轴画在 y 最小值处（图底）
    chart.y_axis.crosses = "min"   # y 轴画在 x 最小值处（图左）

    # 显示坐标轴刻度和标签（"nextTo" = 紧靠轴线，适合两轴均在边缘的情况）
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.tickLblPos = "nextTo"
    chart.y_axis.tickLblPos = "nextTo"
    chart.x_axis.majorTickMark = "out"
    chart.y_axis.majorTickMark = "out"

    # y 轴保留主网格线（后处理设置浅灰虚线）；x 轴无网格线
    from openpyxl.chart.axis import ChartLines
    chart.y_axis.majorGridlines = ChartLines()   # 占位，后处理中设样式
    chart.x_axis.majorGridlines = None           # 无竖向网格线

    # 手动约束绘图区：左留 25%（y 轴标题＋刻度数字＋间距），底留 35%（x 轴刻度＋标题＋图例＋间距）
    # 必须赋给 chart.layout，_write() 会将其同步到 plot_area.layout；
    # 直接赋 chart.plot_area.layout 会被 _write() 中的 self.layout（None）覆盖，导致无效。
    chart.layout = Layout(
        manualLayout=ManualLayout(
            xMode="edge", yMode="edge",
            wMode="edge", hMode="edge",
            x=0.05, y=0.05,
            w=0.97, h=0.90,
        )
    )

    # 放大坐标轴标签字号（best-effort，openpyxl 版本不同可能无效）
    try:
        from openpyxl.drawing.text import (
            RichTextProperties, Paragraph, ParagraphProperties,
            CharacterProperties, ListStyle,
        )

        def _make_txPr(sz_pt: int):
            cp = CharacterProperties(sz=sz_pt * 100)
            pp = ParagraphProperties(defRPr=cp)
            return RichTextProperties(
                bodyPr=None,
                lstStyle=ListStyle(),
                p=[Paragraph(pPr=pp, endParaRPr=cp)],
            )

        chart.y_axis.txPr = _make_txPr(21)
        chart.x_axis.txPr = _make_txPr(21)
    except Exception:
        pass

    # 图例放在 x 轴下方，不与绘图区重叠
    chart.legend.position = "b"
    chart.legend.overlay  = False

    # 图例去重：隐藏 si>0 的 DOW 标记系列（只保留 si=0 的 7 条周一~周日）
    from openpyxl.chart.legend import LegendEntry
    for _hide_idx in range(n_servers + 7, n_servers + n_servers * 7):
        _le = LegendEntry()
        _le.idx = _hide_idx
        _le.delete = True
        chart.legend.legendEntry.append(_le)

    # x 轴数值（开服天数，实际数字）
    x_ref = Reference(ws_chart,
                      min_col=1, max_col=1,
                      min_row=FIRST_DATA_ROW,
                      max_row=FIRST_DATA_ROW + n_days - 1)

    # 每个服务器一条折线（无标记点，形状由星期覆盖层提供）
    for ci, srv in enumerate(servers, start=2):
        y_ref = Reference(ws_chart,
                          min_col=ci, max_col=ci,
                          min_row=FIRST_DATA_ROW,
                          max_row=FIRST_DATA_ROW + n_days - 1)
        ser = Series(y_ref, xvalues=x_ref, title=srv)

        color = _get_server_color(srv)
        ser.graphicalProperties.line.solidFill = color
        ser.graphicalProperties.line.width = 28575   # 2.25 pt（1 pt = 12700 EMU）
        ser.smooth = False
        ser.marker.symbol = "none"

        chart.series.append(ser)

    # n_servers × 7 条标记系列（每服务器每星期，仅标记点无连接线）
    # si=0 的星期系列标题显示在图例（周一~周日）；si>0 的在 XML 后处理中从图例隐藏
    for si, srv in enumerate(servers):
        srv_color = _get_server_color(srv)
        for dj, (dow_label, marker_shape) in enumerate(zip(DOW_LABELS, DOW_MARKERS)):
            mc = MARKER_START_COL + si * 7 + dj
            y_ref = Reference(ws_chart,
                              min_col=mc, max_col=mc,
                              min_row=FIRST_DATA_ROW,
                              max_row=FIRST_DATA_ROW + n_days - 1)
            ser = Series(y_ref, xvalues=x_ref,
                         title=(dow_label if si == 0 else None))

            ser.graphicalProperties.line.noFill = True   # 无连接线
            ser.smooth = False
            ser.marker.symbol = marker_shape
            ser.marker.size   = 8
            if marker_shape in DOW_STROKED:
                # 描边标记（x, plus, dash）：仅设边线色；加 solidFill 会渲染成实心方块
                ser.marker.graphicalProperties.line.solidFill = srv_color
            else:
                # 填充标记：白色填充 + 服务器色边框，保持轻盈感同时保留服务器身份
                ser.marker.graphicalProperties.solidFill = "FFFFFF"
                ser.marker.graphicalProperties.line.solidFill = srv_color

            chart.series.append(ser)

    # 图表锚定在 A4（左上角与 A4 对齐），数据表在图表下方
    ws_chart.add_chart(chart, "A4")

    # ════════════════════════════════════════════════════════════════════════
    # Sheet 2: 数据（原始，未聚合）
    # ════════════════════════════════════════════════════════════════════════
    ws_data = wb.create_sheet("数据")

    DATA_HEADERS = ["商品名", "服务器", "开服天数", "最低价", "均价", "星期"]
    DATA_WIDTHS  = [18, 10, 10, 12, 12, 6]
    for ci, (h, w) in enumerate(zip(DATA_HEADERS, DATA_WIDTHS), start=1):
        c = ws_data.cell(row=1, column=ci, value=h)
        c.fill = fill("2E4057")
        c.font = myfont(bold=True, color="FFFFFF")
        c.alignment = center
        c.border = b_all
        ws_data.column_dimensions[get_column_letter(ci)].width = w
    ws_data.row_dimensions[1].height = 20

    df_sorted = df.sort_values(["category", "item_name", "server_name", "days_open"])
    for ri, (_, row) in enumerate(df_sorted.iterrows(), start=2):
        dow_val = int(row["dow"]) if row["dow"] >= 0 else ""
        for ci, val in enumerate(
            [row["item_name"], row["server_name"], row["days_open"],
             row["min_price"], row["avg_price"], dow_val], start=1
        ):
            c = ws_data.cell(row=ri, column=ci, value=val)
            c.border = b_all
            c.alignment = left_a if ci <= 2 else center
            c.font = myfont()
            if ci in (4, 5):
                c.number_format = "#,##0"
    ws_data.freeze_panes = "A2"

    # ════════════════════════════════════════════════════════════════════════
    # Sheet 3: 商品列表（隐藏，供下拉验证用）
    # ════════════════════════════════════════════════════════════════════════
    ws_items = wb.create_sheet("商品列表")
    for ri, name in enumerate(items, start=1):
        ws_items.cell(row=ri, column=1, value=name)
    ws_items.sheet_state = "hidden"

    # ── 保存 ──────────────────────────────────────────────────────────────
    import shutil
    os.makedirs(ANALYSIS_ROOT, exist_ok=True)
    _out_xlsx = os.path.join(ANALYSIS_ROOT, "price_trend.xlsx")
    _out_xlsm = os.path.join(ANALYSIS_ROOT, "price_trend.xlsm")
    _tmp_xlsx = os.path.join(ANALYSIS_ROOT, "price_trend_tmp.xlsx")

    # 先直接保存为 xlsx（始终有效的 openpyxl 输出）
    wb.save(_out_xlsx)
    # XML 后处理：背景色、网格线、轴线颜色、图例项去重
    _post_process_chart_xml(_out_xlsx, n_servers)

    # VBA 注入：在 xlsx 的副本上操作，避免 Excel COM 修改/锁定原始 xlsx
    # 若注入成功 → 另存为 xlsm，删除 xlsx 和副本
    # 若注入失败 → xlsx 本身完全不受影响，删除副本
    shutil.copy2(_out_xlsx, _tmp_xlsx)
    if _try_inject_vba(_tmp_xlsx, _out_xlsm):
        _saved_name = "price_trend.xlsm"
        for _f in (_out_xlsx, _tmp_xlsx):
            try:
                if os.path.exists(_f):
                    os.remove(_f)
            except Exception:
                pass
    else:
        _saved_name = "price_trend.xlsx"
        for _f in (_tmp_xlsx, _out_xlsm):
            try:
                if os.path.exists(_f):
                    os.remove(_f)
            except Exception:
                pass

    print(f"价格趋势报表已保存：analysis/{_saved_name}"
          f"（{n_items} 件商品，{n_servers} 个服务器，{n_days} 个开服天数）")


def main():
    date_str = prompt_date()
    print(f"\n分析日期：{date_str}")

    print("读取原始数据（当前及历史）…")
    df = load_all_source_csvs(date_str)
    servers = sorted(df["server_name"].unique())
    print(f"  共 {len(df)} 行，涉及服务器：{servers}")

    print("聚合市场快照…")
    market = aggregate_market(df)
    days_list = sorted(market["days_open"].unique())
    print(f"  聚合后：{len(market)} 条记录，开服天数范围 {days_list[0]}~{days_list[-1]} 天")

    # 找出心动服在 date_str 当天的开服天数（需在 build_analysis 前确定）
    xindong_days = None
    today_dir = os.path.join(DATA_ROOT, date_str)
    if os.path.isdir(today_dir):
        for path in glob.glob(os.path.join(today_dir, "*.csv")):
            try:
                tmp = pd.read_csv(path, encoding="utf-8-sig")
                tmp_xd = tmp[tmp["server_name"] == "心动"]
                if not tmp_xd.empty:
                    val = pd.to_numeric(tmp_xd["days_open"].iloc[0], errors="coerce")
                    if not pd.isna(val):
                        xindong_days = int(val)
                        break
            except Exception:
                continue
    if xindong_days is not None:
        print(f"  心动服当前开服天数：{xindong_days} 天，仅以此天数作为买入基准计算套利。")
    else:
        print("  [警告] 未找到心动服当前开服天数，将枚举所有买入天数。")

    print("计算套利分析…")
    result = build_analysis(market, fixed_buy_day=xindong_days)

    out_dir = os.path.join(ANALYSIS_ROOT, date_str)
    os.makedirs(out_dir, exist_ok=True)

    # 完整明细（买入天数已限定为心动当前开服天数）
    out_detail = os.path.join(out_dir, "price_analysis.csv")
    result.to_csv(out_detail, index=False, encoding="utf-8-sig")
    print(f"完整明细已保存：analysis/{date_str}/price_analysis.csv（{len(result)} 行）")

    # 套利机会：每商品保留 min 收益率最高的一条
    tradeable = (result
                 .sort_values("收益率%_min", ascending=False)
                 .drop_duplicates(subset="商品名", keep="first")
                 .reset_index(drop=True))
    out_trade = os.path.join(out_dir, "tradeable_opportunities.csv")
    tradeable.to_csv(out_trade, index=False, encoding="utf-8-sig")
    print(f"套利机会已保存：analysis/{date_str}/tradeable_opportunities.csv（{len(tradeable)} 行）")

    # 精简 xlsx 报表
    export_xlsx(tradeable, out_dir, date_str)
    print(f"精简报表已保存：analysis/{date_str}/top_opportunities.xlsx")

    # 控制台摘要
    if not tradeable.empty:
        print("\n── 套利机会（min 收益率前 20）──")
        cols = ["商品名", "分类", "买入天数", "卖出天数", "天数差",
                "买入价(最低)", "卖出价(最低)", "利润_min", "收益率%_min",
                "买入价(均价)", "卖出价(均价)", "利润_avg", "收益率%_avg"]
        print(tradeable[cols].head(20).to_string(index=False))
    else:
        print("\n暂无套利机会。")

    # 更新历史汇总
    print("\n更新历史汇总报表…")
    build_history_xlsx()

    # 价格趋势折线图报表
    print("\n生成价格趋势报表…")
    build_price_trend_xlsx(df)


if __name__ == "__main__":
    main()
