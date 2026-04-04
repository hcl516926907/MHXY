"""
梦幻西游跨服套利分析 —— 每日运行脚本

用法：
  python daily_analysis.py
  启动后输入要分析的日期（默认今天），程序读取 data/XXXXXX/ 下的 CSV 文件，
  将结果写入 analysis/XXXXXX/。

逻辑：
  以心动服为买入基准（当前开服天数最小），其余服视为"未来"价格参考。
  对每件商品，计算两个场景：
    min 场景：在心动服以 min_price 买入，在目标服以 min_price 卖出
    avg 场景：在心动服以 avg_price 买入，在目标服以 avg_price 卖出
  卖出收入扣除 9% 交易税。
  冻结期规则：
    高级兽诀、高级内丹：购买后 30 天内不可交易
    其余商品：         购买后  7 天内不可交易
  当 (心动开服天数 + 冻结天数) > 目标服开服天数 时，该组合标记为不可交易。

输出（保存在 analysis/XXXXXX/）：
  price_analysis.csv          —— 所有组合（含不可交易）
  tradeable_opportunities.csv —— 仅可交易组合，按 min 收益率降序排列
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

# ── 配置 ──────────────────────────────────────────────────────────────────────
ROOT_DIR    = os.path.dirname(os.path.abspath(__file__))
DATA_ROOT   = os.path.join(ROOT_DIR, "data")
ANALYSIS_ROOT = os.path.join(ROOT_DIR, "analysis")

TAX_RATE = 0.09

# 冻结天数（按商品分类）；不在表中的默认 7 天
FREEZE_DAYS_BY_CATEGORY: dict[str, int] = {
    "高级兽诀": 30,
    "高级内丹": 30,
}
DEFAULT_FREEZE_DAYS = 7

BASE_SERVER = "心动"
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


def load_source_csvs(date_str: str) -> pd.DataFrame:
    """
    读取 data/XXXXXX/ 目录下所有 CSV 文件。
    """
    data_dir = os.path.join(DATA_ROOT, date_str)
    if not os.path.isdir(data_dir):
        raise FileNotFoundError(
            f"未找到目录 data/{date_str}/，请确认已在该日期扫描过数据。"
        )
    paths = glob.glob(os.path.join(data_dir, "*.csv"))
    if not paths:
        raise FileNotFoundError(f"data/{date_str}/ 目录下没有 CSV 文件。")

    frames = []
    for path in paths:
        try:
            df = pd.read_csv(path, encoding="utf-8-sig")
            frames.append(df)
        except Exception as e:
            print(f"  [警告] 跳过文件 {os.path.basename(path)}：{e}")

    if not frames:
        raise ValueError(f"data/{date_str}/ 中所有文件均读取失败。")

    return pd.concat(frames, ignore_index=True)


def build_analysis(df: pd.DataFrame) -> pd.DataFrame:
    """
    核心计算：对每个「心动商品 × 目标服」组合，分别计算 min/avg 两个场景的利润与收益率。

    min 场景：心动 min_price 买入，目标服 min_price 卖出
    avg 场景：心动 avg_price 买入，目标服 avg_price 卖出
    两个场景完全独立，买入价和卖出价均与场景对应。
    """
    df = df.copy()
    df["avg_price"] = pd.to_numeric(df["avg_price"], errors="coerce")
    df["min_price"] = pd.to_numeric(df["min_price"], errors="coerce")
    df["days_open"] = pd.to_numeric(df["days_open"], errors="coerce")

    base   = df[df["server_name"] == BASE_SERVER].copy()
    others = df[df["server_name"] != BASE_SERVER].copy()

    if base.empty:
        raise ValueError(f"数据中未找到基准服务器「{BASE_SERVER}」。")

    base_day = int(base["days_open"].iloc[0])

    records = []
    for _, buy in base.iterrows():
        item        = buy["item_name"]
        category    = buy["category"]
        buy_min     = buy["min_price"]
        buy_avg     = buy["avg_price"]
        freeze_days = FREEZE_DAYS_BY_CATEGORY.get(category, DEFAULT_FREEZE_DAYS)
        earliest_day = base_day + freeze_days

        # 跳过买入价无效的行
        if any(pd.isna(v) for v in [buy_min, buy_avg]):
            continue
        if buy_min <= 0 or buy_avg <= 0:
            continue

        targets = others[others["item_name"] == item]
        for _, sell in targets.iterrows():
            target_server = sell["server_name"]
            target_day    = int(sell["days_open"])
            sell_min      = sell["min_price"]
            sell_avg      = sell["avg_price"]
            sell_count    = pd.to_numeric(sell.get("count", 0), errors="coerce")

            # 跳过卖出价无效的行
            if any(pd.isna(v) for v in [sell_min, sell_avg]):
                continue

            # 目标服采集到的商品数不足 4 件时，挂单稀少，价格不具参考性，跳过
            if pd.isna(sell_count) or sell_count <= 4:
                continue

            tradeable = target_day >= earliest_day
            days_diff = target_day - base_day

            # min 场景：心动 min 买入，目标服 min 卖出
            sell_net_min = sell_min * (1 - TAX_RATE)
            profit_min   = sell_net_min - buy_min
            roi_min      = profit_min / buy_min * 100

            # avg 场景：心动 avg 买入，目标服 avg 卖出
            sell_net_avg = sell_avg * (1 - TAX_RATE)
            profit_avg   = sell_net_avg - buy_avg
            roi_avg      = profit_avg / buy_avg * 100

            records.append({
                "商品名":        item,
                "分类":          category,
                "冻结天数":      freeze_days,
                "基准服":        BASE_SERVER,
                "基准_开服天数": base_day,
                "基准_买入min":  int(buy_min),
                "基准_买入avg":  int(buy_avg),
                "目标服":        target_server,
                "目标_开服天数": target_day,
                "天数差":        days_diff,
                "最早可卖天数":  earliest_day,
                "可交易":        "是" if tradeable else f"否(需≥{earliest_day}天)",
                "目标_卖出min":  int(sell_min),
                "目标_卖出avg":  int(sell_avg),
                "税后收入_min":  round(sell_net_min),
                "税后收入_avg":  round(sell_net_avg),
                "利润_min":      round(profit_min),
                "利润_avg":      round(profit_avg),
                "收益率%_min":   round(roi_min, 1),
                "收益率%_avg":   round(roi_avg, 1),
            })

    result = pd.DataFrame(records)
    result.sort_values(["商品名", "目标_开服天数"], inplace=True, ignore_index=True)
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
        ("买入服",       "基准服"),
        ("买入天数",     "基准_开服天数"),
        ("卖出服",       "目标服"),
        ("卖出天数",     "目标_开服天数"),
        ("天数差",       "天数差"),
        ("买入价(最低)", "基准_买入min"),
        ("卖出价(最低)", "目标_卖出min"),
        ("税后利润",     "利润_min"),
        ("收益率%",      "收益率%_min"),
        ("买入价(均价)", "基准_买入avg"),
        ("卖出价(均价)", "目标_卖出avg"),
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

    # 固定列宽
    COL_WIDTHS = [12, 10, 7, 8, 7, 8, 7, 13, 13, 11, 9, 13, 13, 11, 9]
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
        # 列 1-11 = 最低价组，12-15 = 均价组
        # 在表头加视觉分隔（左边框加粗）
        sep_col = 12  # 均价区起始列
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
                elif src in ("基准_买入min", "目标_卖出min", "税后收入_min",
                             "基准_买入avg", "目标_卖出avg", "税后收入_avg"):
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
        f"▶ 高级兽诀 / 高级内丹（冻结 30 天，仅限天数差≥30 的服务器）",
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


def main():
    date_str = prompt_date()
    print(f"\n分析日期：{date_str}")

    print("读取原始数据…")
    df = load_source_csvs(date_str)
    servers = sorted(df["server_name"].unique())
    print(f"  共 {len(df)} 行，涉及服务器：{servers}")

    print("计算套利分析…")
    result = build_analysis(df)

    out_dir = os.path.join(ANALYSIS_ROOT, date_str)
    os.makedirs(out_dir, exist_ok=True)

    # 完整明细
    out_detail = os.path.join(out_dir, "price_analysis.csv")
    result.to_csv(out_detail, index=False, encoding="utf-8-sig")
    print(f"完整明细已保存：analysis/{date_str}/price_analysis.csv（{len(result)} 行）")

    # 可交易机会
    tradeable = result[result["可交易"] == "是"].copy()
    tradeable.sort_values("收益率%_min", ascending=False, inplace=True, ignore_index=True)
    out_trade = os.path.join(out_dir, "tradeable_opportunities.csv")
    tradeable.to_csv(out_trade, index=False, encoding="utf-8-sig")
    print(f"可交易机会已保存：analysis/{date_str}/tradeable_opportunities.csv（{len(tradeable)} 行）")

    # 精简 xlsx 报表
    xlsx_path = export_xlsx(tradeable, out_dir, date_str)
    print(f"精简报表已保存：analysis/{date_str}/top_opportunities.xlsx")

    # 控制台摘要
    if not tradeable.empty:
        print("\n── 可交易机会（min 收益率前 20）──")
        cols = ["商品名", "分类", "目标服", "天数差",
                "基准_买入min", "目标_卖出min", "利润_min", "收益率%_min",
                "基准_买入avg", "目标_卖出avg", "利润_avg", "收益率%_avg"]
        print(tradeable[cols].head(20).to_string(index=False))
    else:
        print("\n暂无可交易机会。")

    frozen = result[result["可交易"] != "是"]
    print(f"\n── 冻结期内不可交易：{len(frozen)} 条 ──")


if __name__ == "__main__":
    main()
