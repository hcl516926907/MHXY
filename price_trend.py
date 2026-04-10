"""
价格趋势图表模块

职责：
  - 生成 analysis/price_trend.xlsx（或 .xlsm）
  - 包含：图表数据 sheet（商品/价格类型下拉 + 散点图）、数据 sheet、商品列表 sheet（隐藏）
  - 可选：通过 Excel COM 注入 VBA 实现 Y 轴动态缩放（需 win32com）
"""

import os
import shutil

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.layout import Layout, ManualLayout

from paths import ANALYSIS_ROOT

# ── 图表配色常量 ───────────────────────────────────────────────────────────────
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

# 星期相关常量（圆/方/三角/菱形/×/加/五角星 → 周一~周日）
DOW_LABELS  = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
DOW_MARKERS = ["circle", "square", "triangle", "diamond", "x", "plus", "star"]
# 描边标记（只有边线，无实心区域）：x, plus, star
DOW_STROKED: dict = {}
# ─────────────────────────────────────────────────────────────────────────────

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


def _get_server_color(name: str) -> str:
    """返回服务器的固定颜色；未知服务器 fallback 到 _SERIES_COLORS。"""
    return _SERVER_COLORS.get(name, _SERIES_COLORS[abs(hash(name)) % len(_SERIES_COLORS)])


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
    import zipfile
    import io
    import re as _re

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

    # ── 矩阵数据（每行一个开服天数）─────────────────────────────────────────
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
    chart.x_axis.crosses = "min"   # x 轴画在 y 最小值处（图底）
    chart.y_axis.crosses = "min"   # y 轴画在 x 最小值处（图左）

    # 显示坐标轴刻度和标签
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

    # 手动约束绘图区
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
