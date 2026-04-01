import json
import os

DEFAULT_CONFIG = {
    "hotkey": "f9",
    "roi_ratios": {
        "left": 0.26,
        "top": 0.23,
        "right": 0.55,
        "bottom": 0.57,
    },
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
