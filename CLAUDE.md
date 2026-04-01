# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

梦幻西游手游 PC 客户端商城价格扫描工具。用户手动翻页，按 F9 截图 → PaddleOCR 识别 → 解析价格 → 追加 CSV。**不读写内存，不注入进程，纯视觉识别。**

## Running

```bash
# 必须以管理员权限运行（keyboard 全局热键需要）
python main.py
```

- **F9**：扫描当前商城页
- **ESC**：退出程序
- 调试截图输出到 `debug/`（红框标注 ROI 位置 + ROI 裁剪结果）
- 价格数据追加到 `data/prices_YYYYMMDD.csv`

## Architecture

```
main.py       → 热键监听 + scan() 主流程
config.py     → 加载 config.json（缺失键 fallback 到 DEFAULT_CONFIG）
capture.py    → 找游戏窗口 + mss 截图 + ROI 裁剪 + 调试图保存
ocr.py        → PaddleOCR 单例封装（兼容 v2/v3 API）
parser.py     → OCR 文本 → 物品名 + 价格列表
storage.py    → 计算均价 + 追加 CSV
```

