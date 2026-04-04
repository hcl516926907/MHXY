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
- **F10**：自动扫描所有 `target_categories`，完成后自动退出
- **ESC**：退出程序
- 启动时询问是否生成调试截图（默认否）；调试截图输出到 `debug/`
- 价格数据输出到 `data/YYYYMMDD_服务器名.csv`

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
```

## 自动化扫描实现

### 输入模拟

- **点击** (`automator.click`)：`SetCursorPos` + `mouse_event(LEFTDOWN/UP)`，走硬件输入管道
- **拖拽** (`automator.drag`)：逐步移动 + `inertia` 参数控制是否有惯性
  - `inertia=False`（默认）：抬起前静止 0.3s，列表精确停在终点
  - `inertia=True`：立即抬起，游戏引擎触发惯性继续滑动（**不用于回到顶部**，因为截图时动画可能未结束导致坐标错位）
- 坐标使用**客户区坐标**，`client_to_screen()` 转换为屏幕绝对坐标后再操作

### config.json 关键配置项

```json
"target_categories": {"父分类名": ["子商品1", "子商品2", ...]},
"category_roi_ratios": {"left", "top", "right", "bottom"},  // 右侧子商品面板 OCR 区域
"back_roi_ratios":     {"left", "top", "right", "bottom"},  // 左侧父分类导航 OCR 区域
"drag_ratio":          {"x", "y_start", "y_end"},           // 子商品列表翻页拖拽
"scroll_to_top_ratio": {"y_start", "y_end"},                // 回到顶部拖拽（距离更大）
"scroll_to_top_times": 4,       // 回到顶部拖拽次数
"items_per_page": 6,            // 每页条目数，扫满后强制拖拽一次
"scan_wait": 1.0,               // 点击子商品后等待加载
"back_wait": 0.6,               // 点击返回后等待
"drag_wait": 1.0                // 每次拖拽后等待
```

### 自动扫描主循环（auto_scan.py）

```
for 每个父分类:
    OCR 左侧导航面板 → 找到父分类名 → click()
    拖拽回顶部（inertia=False，独立大距离配置）× scroll_to_top_times
    while 还有未扫子商品:
        截图 → OCR 右侧子商品面板
        找出本轮可见目标 → 依次 click → scan() → click 父分类返回
        每扫满 items_per_page 件 → drag() 翻页 → 重新 OCR
        若本轮无目标 → drag() 翻页（最多 auto_scan_max_drags 次）
扫描完成 → os._exit(0)
```

### 商品名匹配规则

- `_norm(text)`：去除 `·`（U+00B7）和 `・`（U+30FB 片假名中点）及空格，用于模糊匹配
- OCR 结果同时以原始名和归一化名入字典，透明支持带点名称（如"汇灵盂·精华"）
- 单字名称（如"毒"）需在 `category_roi_ratios` OCR 中长度过滤 `>= 1` 才可识别

### 空货架处理

`scan()` 返回 `False` 时（未识别到价格），`auto_scan` 调用 `save_na_result()` 在 CSV 写入一行 `avg_price=NA, min_price=NA, count=0`，不跳过该商品。

### CSV 格式

文件名：`YYYYMMDD_服务器名.csv`
列：`timestamp | server_name | open_date | item_name | category | avg_price | min_price | count | prices_raw | days_open`

