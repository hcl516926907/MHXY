"""
Dashboard 生成模块

职责：
  - 将套利分析结果和价格趋势数据汇总为单页 HTML Dashboard
  - 输出 analysis/dashboard.html，可通过 GitHub Pages 发布为公开 URL
"""

import json
import os
import re

import pandas as pd

from paths import ANALYSIS_ROOT, DASHBOARD_HTML
from arbitrage import FREEZE_DAYS_BY_CATEGORY, DEFAULT_FREEZE_DAYS

_SERVERS = ["大吉大利", "天命", "心动", "飞天"]
_SERVER_COLORS = {
    "大吉大利": "#0072B2",
    "天命":     "#E69F00",
    "心动":     "#009E73",
    "飞天":     "#CC3311",
}
_DOW_MAP = {0: "周一", 1: "周二", 2: "周三", 3: "周四", 4: "周五", 5: "周六", 6: "周日"}


# ── 数据整理 ───────────────────────────────────────────────────────────────────

def _build_price_json(df_raw: pd.DataFrame) -> dict:
    """整理价格趋势数据，输出嵌套 JSON 结构。"""
    df = df_raw.copy()
    for col in ("avg_price", "min_price", "days_open", "count"):
        df[col] = pd.to_numeric(df[col], errors="coerce")
    # 与 price_trend.py 保持一致：过滤无效行
    df = df[
        df["avg_price"].notna() & (df["avg_price"] > 0) &
        df["min_price"].notna() & (df["min_price"] > 0) &
        df["days_open"].notna() &
        df["count"].notna() & (df["count"] >= 1)
    ].copy()
    df["days_open"] = df["days_open"].astype(int)

    # 用 format="mixed" 兼容多种时间戳格式（如 "2026/4/3 14:54" vs "2026-04-02 19:27:28"）
    df["_ts"] = pd.to_datetime(df["timestamp"], errors="coerce", format="mixed")
    df["dow"] = df["_ts"].dt.dayofweek.map(_DOW_MAP).fillna("周一")

    # 同一 (商品, 服务器, 开服天数) 保留最新一条
    df = (df.sort_values("_ts")
            .drop_duplicates(subset=["item_name", "server_name", "days_open"], keep="last")
            .drop(columns=["_ts"]))

    items = sorted(df["item_name"].unique())
    data = {}
    freeze = {}  # item_name -> 冻结天数
    for item in items:
        idf = df[df["item_name"] == item]
        # 取该商品的分类，查冻结天数
        category = idf["category"].iloc[0] if "category" in idf.columns else ""
        freeze[item] = FREEZE_DAYS_BY_CATEGORY.get(category, DEFAULT_FREEZE_DAYS)
        data[item] = {}
        for server in _SERVERS:
            sdf = idf[idf["server_name"] == server].sort_values("days_open")
            data[item][server] = [
                {
                    "x":   int(r["days_open"]),
                    "min": int(r["min_price"]),
                    "avg": int(r["avg_price"]),
                    "dow": r["dow"],
                }
                for _, r in sdf.iterrows()
            ]

    return {"items": items, "servers": _SERVERS, "colors": _SERVER_COLORS, "data": data, "freeze": freeze}


def _v(row, col, cast=float):
    """安全取值并类型转换，失败返回 0。"""
    v = row.get(col, None) if hasattr(row, "get") else getattr(row, col, None)
    try:
        return cast(v)
    except (TypeError, ValueError):
        return 0


def _build_history_json() -> dict:
    """扫描历史 tradeable_opportunities.csv，提取每日 top5×2。"""
    all_rows = []
    dates = []

    if not os.path.isdir(ANALYSIS_ROOT):
        return {"dates": [], "records": []}

    for entry in sorted(os.listdir(ANALYSIS_ROOT), reverse=True):
        if not re.fullmatch(r"\d{8}", entry):
            continue
        csv_path = os.path.join(ANALYSIS_ROOT, entry, "tradeable_opportunities.csv")
        if not os.path.isfile(csv_path):
            continue
        try:
            df = pd.read_csv(csv_path, encoding="utf-8-sig")
        except Exception as e:
            print(f"  [dashboard 警告] 跳过 {entry}: {e}")
            continue

        dates.append(entry)
        for freeze in [30, 7]:
            subset = (df[df["冻结天数"] == freeze]
                      .sort_values("收益率%_min", ascending=False)
                      .head(5))
            for rank, (_, row) in enumerate(subset.iterrows(), start=1):
                all_rows.append({
                    "date":       entry,
                    "freeze":     freeze,
                    "rank":       rank,
                    "item":       str(row.get("商品名", "")),
                    "category":   str(row.get("分类", "")),
                    "buy_day":    _v(row, "买入天数", int),
                    "sell_day":   _v(row, "卖出天数", int),
                    "diff":       _v(row, "天数差", int),
                    "buy_min":    _v(row, "买入价(最低)"),
                    "sell_min":   _v(row, "卖出价(最低)"),
                    "profit_min": _v(row, "利润_min"),
                    "roi_min":    _v(row, "收益率%_min"),
                    "buy_avg":    _v(row, "买入价(均价)"),
                    "sell_avg":   _v(row, "卖出价(均价)"),
                    "profit_avg": _v(row, "利润_avg"),
                    "roi_avg":    _v(row, "收益率%_avg"),
                })

    return {"dates": dates, "records": all_rows}


# ── 主入口 ─────────────────────────────────────────────────────────────────────

def build_dashboard_html(df_raw: pd.DataFrame, date_str: str):
    """
    生成 analysis/dashboard.html。

    参数：
      df_raw   — 原始价格 DataFrame（来自 load_all_source_csvs）
      date_str — 分析日期 "YYYYMMDD"
    """
    print("  整理价格趋势数据…")
    price_data = _build_price_json(df_raw)
    print("  整理历史套利数据…")
    history = _build_history_json()

    html = _TEMPLATE
    html = html.replace("__PRICE_DATA__", json.dumps(price_data, ensure_ascii=False))
    html = html.replace("__HISTORY__",    json.dumps(history,    ensure_ascii=False))
    html = html.replace("__DATE__",       date_str)

    os.makedirs(ANALYSIS_ROOT, exist_ok=True)
    with open(DASHBOARD_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  Dashboard 已生成：analysis/dashboard.html")


# ── HTML 模板 ──────────────────────────────────────────────────────────────────

_TEMPLATE = """\
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>梦幻西游 套利 Dashboard</title>
<script src="https://cdn.plot.ly/plotly-latest.min.js" charset="utf-8"></script>
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
       background: #f6f8fa; color: #24292f; font-size: 14px; }

/* ── 顶部栏 ── */
header { background: #24292f; color: #fff; padding: 12px 24px;
         display: flex; align-items: center; justify-content: space-between; }
header h1 { font-size: 16px; font-weight: 600; }
header .meta { font-size: 12px; color: #8b949e; }

/* ── Tab 导航 ── */
.tabs { background: #fff; border-bottom: 1px solid #d0d7de;
        padding: 0 24px; display: flex; }
.tab-btn { background: none; border: none; cursor: pointer; padding: 12px 16px;
           font-size: 14px; color: #57606a; border-bottom: 2px solid transparent;
           transition: color .15s; }
.tab-btn:hover { color: #24292f; }
.tab-btn.active { color: #24292f; border-bottom-color: #fd8c73; font-weight: 600; }

/* ── Tab 面板 ── */
.tab-panel { display: none; padding: 24px; }
.tab-panel.active { display: block; }

/* ── 控件行 ── */
.controls { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin-bottom: 20px; }
.controls label { font-weight: 500; color: #57606a; }
select, input[type=text] {
  border: 1px solid #d0d7de; border-radius: 6px; padding: 5px 10px;
  font-size: 14px; background: #fff; color: #24292f; }
select:focus, input[type=text]:focus { outline: 2px solid #0969da; outline-offset: -1px; }
.btn-group { display: flex; }
.btn-group button {
  border: 1px solid #d0d7de; background: #fff; padding: 5px 14px;
  cursor: pointer; font-size: 13px; color: #24292f; transition: background .12s; }
.btn-group button:first-child { border-radius: 6px 0 0 6px; }
.btn-group button:last-child  { border-radius: 0 6px 6px 0; }
.btn-group button:not(:first-child) { border-left: none; }
.btn-group button.active { background: #0969da; color: #fff; border-color: #0969da; }
.btn-group button:hover:not(.active) { background: #f6f8fa; }

/* ── 双表格区域 ── */
.two-tables { display: block; }
.freeze-section { margin-bottom: 28px; }
.freeze-section h3 {
  font-size: 13px; font-weight: 600; padding: 6px 10px; margin-bottom: 8px;
  border-radius: 4px; }
.freeze-section.f30 h3 { background: #fff8e1; color: #7d5800; border-left: 4px solid #e69f00; }
.freeze-section.f7  h3 { background: #e3f2fd; color: #0d47a1; border-left: 4px solid #0072b2; }

/* ── 表格 ── */
.tbl-wrap { overflow-x: auto; border-radius: 6px; border: 1px solid #d0d7de; }
table { width: 100%; border-collapse: collapse; background: #fff; }
thead th {
  background: #f6f8fa; font-weight: 600; padding: 8px 12px;
  text-align: right; white-space: nowrap;
  border-bottom: 1px solid #d0d7de; }
thead th:first-child, thead th:nth-child(2), thead th:nth-child(3) { text-align: left; }
tbody tr:nth-child(even) { background: #f6f8fa; }
tbody tr:hover { background: #ddf4ff; }
td { padding: 7px 12px; text-align: right; white-space: nowrap;
     border-bottom: 1px solid #eaeef2; }
td:first-child, td:nth-child(2), td:nth-child(3) { text-align: left; }
.profit-pos { color: #1a7f37; font-weight: 600; }
.profit-neg { color: #cf222e; }

/* ── 图表区 ── */
.chart-legend-bar {
  display: flex; align-items: center; gap: 20px; flex-wrap: wrap;
  padding: 7px 12px; background: #fff;
  border: 1px solid #d0d7de; border-bottom: none;
  border-radius: 6px 6px 0 0; }
.srv-item { display: flex; align-items: center; gap: 5px; font-size: 13px; }
.srv-line { width: 22px; height: 3px; border-radius: 2px; display: inline-block; }
.legend-sep { width: 1px; height: 16px; background: #d0d7de; margin: 0 4px; }
.dow-items { display: flex; gap: 12px; font-size: 13px; color: #57606a; align-items: center; }
.dow-items span { display: inline-flex; align-items: center; gap: 3px; }
.dow-sym { display: inline-flex; align-items: center; justify-content: center; width: 16px; }
#trend-chart { width: 100%; height: 520px; background: #fff;
               border: 1px solid #d0d7de; border-radius: 0 0 6px 6px; }

/* ── 商品自动完成 ── */
.item-autocomplete { position: relative; display: inline-block; }
.ac-dropdown {
  display: none; position: absolute; top: 100%; left: 0;
  min-width: 200px; max-height: 240px; overflow-y: auto;
  background: #fff; border: 1px solid #d0d7de; border-radius: 0 0 6px 6px;
  box-shadow: 0 4px 12px rgba(0,0,0,.12); z-index: 100; }
.ac-item { padding: 6px 12px; cursor: pointer; font-size: 14px; white-space: nowrap; }
.ac-item:hover { background: #ddf4ff; }

/* ── 空状态 ── */
.empty-cell { text-align: center !important; padding: 30px !important;
              color: #8b949e; font-style: italic; }
</style>
</head>
<body>

<header>
  <h1>梦幻西游 跨服套利 Dashboard</h1>
  <span class="meta">数据更新：__DATE__</span>
</header>

<nav class="tabs">
  <button class="tab-btn active" data-tab="arb">套利机会总览</button>
  <button class="tab-btn" data-tab="trend">价格趋势</button>
</nav>

<!-- ── Tab 1: 套利机会总览 ── -->
<div id="tab-arb" class="tab-panel active">
  <div class="controls">
    <label>日期：</label>
    <select id="arb-date" onchange="renderArbitrage()" style="min-width:120px"></select>
  </div>
  <div class="two-tables">
    <!-- 30天冻结 -->
    <div class="freeze-section f30">
      <h3>30天冻结（Top 5）</h3>
      <div class="tbl-wrap">
        <table>
          <thead><tr>
            <th>排名</th>
            <th>商品名</th>
            <th>分类</th>
            <th>买入天数</th>
            <th>卖出天数</th>
            <th>天数差</th>
            <th>买入价(最低)</th>
            <th>卖出价(最低)</th>
            <th>利润(最低)</th>
            <th>收益率%(最低)</th>
            <th>买入价(均价)</th>
            <th>卖出价(均价)</th>
            <th>利润(均价)</th>
            <th>收益率%(均价)</th>
          </tr></thead>
          <tbody id="arb-30-body"></tbody>
        </table>
      </div>
    </div>
    <!-- 7天冻结 -->
    <div class="freeze-section f7">
      <h3>7天冻结（Top 5）</h3>
      <div class="tbl-wrap">
        <table>
          <thead><tr>
            <th>排名</th>
            <th>商品名</th>
            <th>分类</th>
            <th>买入天数</th>
            <th>卖出天数</th>
            <th>天数差</th>
            <th>买入价(最低)</th>
            <th>卖出价(最低)</th>
            <th>利润(最低)</th>
            <th>收益率%(最低)</th>
            <th>买入价(均价)</th>
            <th>卖出价(均价)</th>
            <th>利润(均价)</th>
            <th>收益率%(均价)</th>
          </tr></thead>
          <tbody id="arb-7-body"></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<!-- ── Tab 2: 价格趋势 ── -->
<div id="tab-trend" class="tab-panel">
  <div class="controls">
    <label>商品：</label>
    <div class="item-autocomplete" id="item-ac">
      <input type="text" id="item-search" placeholder="搜索或选择商品…" autocomplete="off" style="width:200px">
      <div class="ac-dropdown" id="item-dropdown"></div>
    </div>
    <label>价格：</label>
    <div class="btn-group">
      <button id="price-min-btn" class="active" onclick="setPriceType('min')">最低价</button>
      <button id="price-avg-btn" onclick="setPriceType('avg')">均价</button>
    </div>
  </div>
  <div class="chart-legend-bar">
    <span class="srv-item"><span class="srv-line" style="background:#0072B2"></span>大吉大利</span>
    <span class="srv-item"><span class="srv-line" style="background:#E69F00"></span>天命</span>
    <span class="srv-item"><span class="srv-line" style="background:#009E73"></span>心动</span>
    <span class="srv-item"><span class="srv-line" style="background:#CC3311"></span>飞天</span>
    <span class="legend-sep"></span>
    <span class="srv-item">
      <span style="display:inline-block;width:18px;height:14px;background:rgba(180,180,180,0.45);border:1px solid #aaa;border-radius:2px;vertical-align:middle"></span>
      冻结期（无法出售）
    </span>
    <span class="legend-sep"></span>
    <div class="dow-items">
      <span><span class="dow-sym" style="font-size:20px">&#9679;</span>周一</span>
      <span><span class="dow-sym">&#9632;</span>周二</span>
      <span><span class="dow-sym">&#9650;</span>周三</span>
      <span><span class="dow-sym">&#9670;</span>周四</span>
      <span><svg width="12" height="12" viewBox="0 0 14 14" style="vertical-align:middle;margin-right:3px"><line x1="2" y1="2" x2="12" y2="12" stroke="#57606a" stroke-width="3.5" stroke-linecap="round"/><line x1="12" y1="2" x2="2" y2="12" stroke="#57606a" stroke-width="3.5" stroke-linecap="round"/></svg>周五</span>
      <span><span class="dow-sym">&#10010;</span>周六</span>
      <span><span class="dow-sym">&#9733;</span>周日</span>
    </div>
  </div>
  <div id="trend-chart"></div>
</div>

<script>
// ── 数据 ──────────────────────────────────────────────────────────────────────
const DATA_PRICE   = __PRICE_DATA__;
const DATA_HISTORY = __HISTORY__;

// ── Tab 切换 ──────────────────────────────────────────────────────────────────
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
    if (btn.dataset.tab === 'trend') updateChart();
  });
});

// ── 工具函数 ──────────────────────────────────────────────────────────────────
function fmt(n)    { return (n == null || isNaN(n)) ? '\u2014' : Math.round(n).toLocaleString('zh-CN'); }
function fmtPct(n) { return (n == null || isNaN(n)) ? '\u2014' : n.toFixed(2) + '%'; }
function profitCls(v) { return v > 0 ? 'profit-pos' : v < 0 ? 'profit-neg' : ''; }
function fmtDate(d) {
  if (!d || d.length !== 8) return d;
  return d.slice(0,4) + '-' + d.slice(4,6) + '-' + d.slice(6,8);
}

// ── 套利机会总览 ──────────────────────────────────────────────────────────────
function renderFreezeTable(tbodyId, rows) {
  const tbody = document.getElementById(tbodyId);
  if (rows.length === 0) {
    tbody.innerHTML = '<tr><td colspan="14" class="empty-cell">\u6682\u65e0\u6570\u636e</td></tr>';
    return;
  }
  tbody.innerHTML = rows.map(r => `
    <tr>
      <td>${r.rank}</td>
      <td>${r.item}</td>
      <td>${r.category}</td>
      <td>${r.buy_day}</td>
      <td>${r.sell_day}</td>
      <td>${r.diff}</td>
      <td>${fmt(r.buy_min)}</td>
      <td>${fmt(r.sell_min)}</td>
      <td class="${profitCls(r.profit_min)}">${fmt(r.profit_min)}</td>
      <td class="${profitCls(r.profit_min)}">${fmtPct(r.roi_min)}</td>
      <td>${fmt(r.buy_avg)}</td>
      <td>${fmt(r.sell_avg)}</td>
      <td class="${profitCls(r.profit_avg)}">${fmt(r.profit_avg)}</td>
      <td class="${profitCls(r.profit_avg)}">${fmtPct(r.roi_avg)}</td>
    </tr>`).join('');
}

function renderArbitrage() {
  const selDate = document.getElementById('arb-date').value;
  const records = DATA_HISTORY.records.filter(r => r.date === selDate);
  const f30 = records.filter(r => r.freeze === 30).sort((a, b) => a.rank - b.rank);
  const f7  = records.filter(r => r.freeze === 7).sort((a, b) => a.rank - b.rank);
  renderFreezeTable('arb-30-body', f30);
  renderFreezeTable('arb-7-body',  f7);
}

// ── 价格趋势 ──────────────────────────────────────────────────────────────────
let priceType = 'min';
const allItems = DATA_PRICE.items || [];
let selectedItem = allItems.length > 0 ? allItems[0] : '';

const DOW_SYMBOLS = {
  '\u5468\u4e00': 'circle',      '\u5468\u4e8c': 'square',
  '\u5468\u4e09': 'triangle-up', '\u5468\u56db': 'diamond',
  '\u5468\u4e94': 'x',           '\u5468\u516d': 'cross',
  '\u5468\u65e5': 'star'
};

// ── 商品自动完成 ──────────────────────────────────────────────────────────────
(function setupAutocomplete() {
  const input    = document.getElementById('item-search');
  const dropdown = document.getElementById('item-dropdown');

  function showDropdown(items) {
    if (items.length === 0) { dropdown.style.display = 'none'; return; }
    dropdown._items = items;
    dropdown.innerHTML = items.map((it, i) =>
      `<div class="ac-item" data-idx="${i}">${it}</div>`
    ).join('');
    dropdown.style.display = 'block';
  }

  function hideDropdown() { dropdown.style.display = 'none'; }

  function getFiltered() {
    const q = input.value.trim();
    return q ? allItems.filter(it => it.includes(q)) : allItems;
  }

  input.addEventListener('focus', () => showDropdown(getFiltered()));
  input.addEventListener('input', () => showDropdown(getFiltered()));

  dropdown.addEventListener('mousedown', e => {
    // mousedown fires before blur, so we can safely read the item
    const el = e.target.closest('.ac-item');
    if (!el) return;
    e.preventDefault();   // prevent input from losing focus and hiding dropdown
    const item = dropdown._items[parseInt(el.dataset.idx)];
    selectedItem = item;
    input.value  = item;
    hideDropdown();
    updateChart();
  });

  document.addEventListener('click', e => {
    if (!document.getElementById('item-ac').contains(e.target)) hideDropdown();
  });
})();

function setPriceType(t) {
  priceType = t;
  document.getElementById('price-min-btn').classList.toggle('active', t === 'min');
  document.getElementById('price-avg-btn').classList.toggle('active', t === 'avg');
  updateChart();
}

function updateChart() {
  const item = selectedItem;
  if (!item || !DATA_PRICE.data[item]) { Plotly.purge('trend-chart'); return; }

  const traces = [];
  for (const server of DATA_PRICE.servers) {
    const pts = DATA_PRICE.data[item][server] || [];
    if (pts.length === 0) continue;
    traces.push({
      x:    pts.map(p => p.x),
      y:    pts.map(p => priceType === 'min' ? p.min : p.avg),
      mode: 'lines+markers',
      name: server,
      line:   { color: DATA_PRICE.colors[server], width: 2 },
      marker: {
        symbol: pts.map(p => DOW_SYMBOLS[p.dow] || 'circle'),
        size:   pts.map(p => ({'周三':15,'周日':15})[p.dow] || 12),
        color:  DATA_PRICE.colors[server],
        line:   { width: 1, color: '#fff' }
      },
      hovertemplate: server + '<br>\u5929\u6570: %{x}<br>\u4ef7\u683c: %{y:,.0f}<extra></extra>'
    });
  }

  // 冻结期阴影：取心动服最新天数作为起点
  const xindongPts = DATA_PRICE.data[item]['\u5fc3\u52a8'] || [];
  const shapes = [];
  if (xindongPts.length > 0) {
    const x0 = Math.max(...xindongPts.map(p => p.x));
    const freeze = DATA_PRICE.freeze[item] || 7;
    shapes.push({
      type: 'rect', xref: 'x', yref: 'paper',
      x0: x0, x1: x0 + freeze,
      y0: 0,  y1: 1,
      fillcolor: 'rgba(180,180,180,0.25)',
      line: { width: 0 }
    });
  }

  const layout = {
    xaxis: { title: '\u670d\u52a1\u5668\u5929\u6570', gridcolor: '#eaeef2', zeroline: false },
    yaxis: { title: { text: '\u4ef7\u683c', standoff: 20 }, tickformat: ',.0f', gridcolor: '#eaeef2', zeroline: false },
    showlegend: false,
    shapes: shapes,
    margin: { t: 20, r: 20, b: 60, l: 90 },
    paper_bgcolor: '#fff',
    plot_bgcolor:  '#fff',
    hovermode: 'closest',
  };
  Plotly.newPlot('trend-chart', traces, layout, { responsive: true, displayModeBar: false });
}

// ── 初始化 ────────────────────────────────────────────────────────────────────
(function init() {
  // 套利机会总览 — 填充日期下拉
  const arbSel = document.getElementById('arb-date');
  if (DATA_HISTORY.dates.length > 0) {
    arbSel.innerHTML = DATA_HISTORY.dates
      .map(d => `<option value="${d}">${fmtDate(d)}</option>`).join('');
    arbSel.value = DATA_HISTORY.dates[0];
    renderArbitrage();
  } else {
    arbSel.innerHTML = '<option>\u6682\u65e0\u5386\u53f2\u6570\u636e</option>';
  }

  // 价格趋势 — 初始化搜索框（不在隐藏 tab 里调用 updateChart，避免 Plotly 尺寸错误）
  if (selectedItem) {
    document.getElementById('item-search').value = selectedItem;
  }
})();
</script>
</body>
</html>
"""
