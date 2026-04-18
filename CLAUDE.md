# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

梦幻西游手游 PC 客户端商城价格扫描工具。**不读写内存，不注入进程，纯视觉识别。**
通过扫描心动服、飞天服、大吉大利服、天命服（开服顺序为天命→大吉大利→飞天→心动）的商品价格，进行套利分析，从而在心动服低价购入商品并在后期高价售出。服务器开服间隔为14天。

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
