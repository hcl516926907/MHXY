# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

梦幻西游手游 PC 客户端商城价格扫描工具。**不读写内存，不注入进程，纯视觉识别。**
通过扫描心动服、飞天服、大吉大利服、天命服、来财服、红颜服（开服顺序为红颜→来财→天命→大吉大利→飞天→心动）的商品价格，进行套利分析，从而在心动服低价购入商品并在后期高价售出。服务器开服间隔为14天。

商品分为珍品和非珍品。珍品包含高级兽诀、高级内丹，非珍品包含低级兽诀、奇珍异宝、法宝和低级内丹。珍品有30天的冻结期，非珍品有7天冻结期。

## Running

```bash
# 必须以管理员权限运行（keyboard 全局热键需要）
python main.py
```

## Architecture

```
main.py       → 热键监听 + scan() 主流程 + 自动扫描入口
config.py     → 加载 config.json（缺失键 fallback 到 DEFAULT_CONFIG）
capture.py    → 找游戏窗口 + mss 截图 + ROI 裁剪 + 调试图保存
                set_debug_enabled(bool) 控制是否生成调试图
ocr.py        → PaddleOCR 单例封装（兼容 v2/v3 API）
parser.py     → OCR 文本 → 物品名 + 价格列表（支持单字名称如"毒"）
storage.py    → 计算均价/最低价 + 追加 CSV；save_na_result() 写空货架记录
automator.py  → Win32 输入封装：click() / drag()
auto_scan.py  → 自动扫描主循环
daily_analysis.py → 计算每日套利机会并更新dashboard
```

## 套利计算逻辑

### 利润率（arbitrage.py）

```
买入价  = 心动服在买入天数的 min_price
卖出价  = 卖出天数对应的跨服聚合最低价（所有服务器该开服天数最低价的最小值）
税后收入 = 卖出价 × (1 - 9%)
利润    = 税后收入 - 买入价
收益率% = 利润 / 买入价 × 100
```

条件：卖出天数 - 买入天数 > 冻结天数（珍品30天，非珍品7天）

### 各服增长率（dashboard.py JS）

**目的**：在卖出日那一天，各服务器的近期价格涨幅（判断市场热度）。

**计算步骤**：

1. **确定卖出自然日**：通过 `findReferenceServer` 找到在 sell_day 有精确数据的服务器（优先最新开服），结合该服的 open_date 反推卖出天数对应的公历日期。
2. **换算各服开服天数**：用自然日减去各服 open_date 得到各服在该日的开服天数 `t`。
3. **计算增长率**：`(min(t) - min(t-2)) / min(t-2) × 100`
   - `min(t)` = 该服在第 t 天的精确最低价（无则跳过）
   - `min(t-2)` = 该服在第 t-2 天或之前最近一天的最低价
4. **展示顺序**：心动 → 飞天 → 大吉大利 → 天命（按开服时间由晚到早）

**注意**：open_date 须为 `YYYY-MM-DD` 格式（破折号），否则 JS `new Date()` 在 CST 时区解析会差8小时导致日期偏移。`_build_price_json` 中已对 `/` 做替换。
