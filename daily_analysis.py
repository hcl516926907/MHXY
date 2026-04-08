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
from openpyxl.worksheet.datavalidation import DataValidation

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


# ── 历史汇总相关常量 ──────────────────────────────────────────────────────────
# 历史数据 sheet 列头（按顺序）
_HIST_HEADERS = [
    "日期", "冻结天数", "排名",
    "商品名", "分类", "买入服", "买入天数",
    "卖出服", "卖出天数", "天数差",
    "买入价(最低)", "卖出价(最低)", "税后利润", "收益率%",
    "买入价(均价)", "卖出价(均价)", "税后利润(均)", "收益率%(均)",
]
# tradeable_opportunities.csv 中对应列（前 3 个由代码填充）
_HIST_SRC = [
    None, None, None,
    "商品名", "分类", "基准服", "基准_开服天数",
    "目标服", "目标_开服天数", "天数差",
    "基准_买入min", "目标_卖出min", "利润_min", "收益率%_min",
    "基准_买入avg", "目标_卖出avg", "利润_avg", "收益率%_avg",
]
# Sheet 1 展示列（去掉日期/冻结天数/排名），及对应 历史数据 sheet 列字母
_DISP_HEADERS = _HIST_HEADERS[3:]   # 15 列
_HIST_DATA_COLS = [get_column_letter(i) for i in range(4, 4 + len(_DISP_HEADERS))]  # D~R


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

    n_disp = len(_DISP_HEADERS)   # 15
    COL_WIDTHS = [12, 10, 7, 8, 7, 8, 7, 13, 13, 11, 9, 13, 13, 11, 9]
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

    sep_col = 12   # 均价区起始（1-based）

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
    HIST_COL_WIDTHS = [10, 8, 5, 12, 10, 7, 8, 7, 8, 7, 13, 13, 11, 9, 13, 13, 11, 9]
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

    # 更新历史汇总
    print("\n更新历史汇总报表…")
    build_history_xlsx()


if __name__ == "__main__":
    main()
