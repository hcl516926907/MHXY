import json
import os

DEFAULT_CONFIG = {
    "hotkey": "f9",
    "auto_scan_hotkey": "f10",
    # ROI for the price list (used by F9 manual scan)
    "roi_ratios": {
        "left": 0.26,
        "top": 0.23,
        "right": 0.55,
        "bottom": 0.57,
    },
    # ROI for the category selection grid (right panel) — used by auto-scan OCR
    # Adjust these ratios if the detected names are wrong or missing
    "category_roi_ratios": {
        "left": 0.27,
        "top": 0.18,
        "right": 0.97,
        "bottom": 0.83,
    },
    # ROI for the LEFT navigation panel (parent categories like 高级兽诀, 低级兽诀).
    # Used after scanning a sub-item to find the parent category button and click back.
    # Calibrate if OCR still can't find the parent category name.
    "back_roi_ratios": {"left": 0.12, "top": 0.18, "right": 0.22, "bottom": 0.55},
    # Fallback back-button ratio — only used when back_roi_ratios OCR also fails.
    "back_button_ratio": {"x": 0.18, "y": 0.49},
    # Drag parameters for scrolling the category grid (all ratios of window size)
    "drag_ratio": {"x": 0.62, "y_start": 0.70, "y_end": 0.28},
    # Wait times (seconds)
    "scan_wait": 1.0,   # after clicking a category item, before scanning
    "back_wait": 0.6,   # after clicking back, before next iteration
    "drag_wait": 1.0,   # after dragging, before re-OCR
    # Maximum drag attempts before giving up on unfound items
    "auto_scan_max_drags": 5,
    # Dict of {父分类: [子商品列表]}，名称必须与游戏内文字完全一致
    # 例：{"高级兽诀": ["高级必杀", "高级连击"], "低级兽诀": ["低级必杀"]}
    "target_categories": {},
    "ocr_confidence_threshold": 0.7,
    "use_gpu": True,
    "output_dir": "data",
}

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


def load_config() -> dict:
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            config = json.load(f)
        for k, v in DEFAULT_CONFIG.items():
            config.setdefault(k, v)
        return config
    return DEFAULT_CONFIG.copy()


def save_config(config: dict):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
