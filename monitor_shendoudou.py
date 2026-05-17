# ── 必须在所有其他 import 之前设置 ─────────────────────────────────────────
import os
import sys
import types
import ctypes

if "modelscope" not in sys.modules:
    _ms_stub = types.ModuleType("modelscope")
    sys.modules["modelscope"] = _ms_stub

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass

os.environ["FLAGS_use_mkldnn"] = "0"
os.environ["FLAGS_enable_pir_api"] = "0"
# ────────────────────────────────────────────────────────────────────────────

import time
from datetime import datetime
from typing import Optional

import requests

from config import load_config
from automator import get_game_hwnd, click
import cv2
from capture import find_game_window, capture_window, crop_item_grid, set_debug_enabled
from ocr import recognize, get_ocr
from parser import parse_page
from auto_scan import _click_parent_via_ocr, _find_categories_on_screen, _click_back_via_ocr, _norm


def send_wechat_alert(sckey: str, title: str, desp: str) -> bool:
    url = f"https://sctapi.ftqq.com/{sckey}.send"
    try:
        resp = requests.get(url, params={"title": title, "desp": desp}, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        if data.get("code") == 0:
            print(f"[通知] 已发送 WeChat 推送: {title}")
            return True
        else:
            print(f"[通知] Server酱返回错误: {data}")
            return False
    except requests.RequestException as e:
        print(f"[通知] 推送失败: {e}")
        return False


def navigate_to_shendoudou(hwnd: int, config: dict, item: str, parent: str) -> bool:
    try:
        _click_parent_via_ocr(hwnd, parent, config)
        time.sleep(config.get("back_wait", 0.6))

        rect = find_game_window()
        if rect is None:
            return False
        full_img = capture_window(rect)
        visible = _find_categories_on_screen(full_img, config)

        # 精确匹配或归一化匹配
        key = item if item in visible else _norm(item)
        if key not in visible:
            print(f"    [导航] 未在子商品列表中找到「{item}」，识别到: {list(visible.keys())}")
            return False

        cx, cy = visible[key]
        click(hwnd, cx, cy)
        time.sleep(config.get("scan_wait", 1.0))
        return True
    except Exception as e:
        print(f"[导航错误] {e}")
        return False


def capture_and_parse_price(config: dict) -> Optional[int]:
    min_valid = config.get("monitor", {}).get("min_valid_price", 40000)

    rect = find_game_window()
    if rect is None:
        return None
    full_img = capture_window(rect)
    grid_img = crop_item_grid(full_img, config["roi_ratios"])
    texts = recognize(
        grid_img,
        use_gpu=config.get("use_gpu", False),
        confidence_threshold=config.get("ocr_confidence_threshold", 0.7),
    )
    _, prices = parse_page(texts)
    if not prices:
        return None

    raw_min = min(prices)
    if raw_min < min_valid:
        debug_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug")
        os.makedirs(debug_dir, exist_ok=True)
        ts = datetime.now().strftime("%H%M%S")
        path = os.path.join(debug_dir, f"monitor_price_roi_suspicious_{ts}.png")
        ok, buf = cv2.imencode(".png", grid_img)
        if ok:
            with open(path, "wb") as f:
                f.write(buf.tobytes())
        print(f"    [调试] 可疑价格 {raw_min:,}，裁剪图已保存: {path}")

    valid_prices = [p for p in prices if p >= min_valid]
    return min(valid_prices) if valid_prices else None


def navigate_back(hwnd: int, config: dict, parent: str) -> None:
    try:
        _click_back_via_ocr(hwnd, parent, config)
        time.sleep(config.get("back_wait", 0.6))
    except Exception as e:
        print(f"[返回错误] {e}（下轮将重新导航）")


def monitor_loop(config: dict) -> None:
    mon = config.get("monitor", {})
    item      = mon.get("item", "神兜兜")
    parent    = mon.get("parent_category", "奇珍异宝")
    threshold = mon.get("price_threshold", 75000)
    interval  = mon.get("interval_seconds", 10)
    cooldown  = mon.get("alert_cooldown_seconds", 300)
    sckey     = mon.get("serverchan_sckey", "")

    set_debug_enabled(False)  # 禁用 auto_scan 内部的常规调试截图

    print(f"[神兜兜监控] 阈值: {threshold:,} | 间隔: {interval}s | 冷却: {cooldown}s | "
          f"WeChat推送: {'已配置' if sckey and 'SendKey' not in sckey else '未配置'}")
    print("按 Ctrl+C 停止\n")

    # 预热 OCR 模型，避免第一轮卡顿
    print("[初始化] 加载 OCR 模型...")
    get_ocr(use_gpu=config.get("use_gpu", False))
    print("[初始化] 完成，开始监控\n")

    last_alert_time: float = 0.0

    try:
        while True:
            cycle_start = time.time()
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            hwnd = get_game_hwnd()
            if hwnd is None:
                print(f"[{now_str}] 未找到游戏窗口，跳过本轮")
            else:
                ok = navigate_to_shendoudou(hwnd, config, item, parent)
                if ok:
                    min_price = capture_and_parse_price(config)
                    if min_price is not None:
                        flag = " *** 低价！" if min_price < threshold else ""
                        print(f"[{now_str}] {item} 最低价: {min_price:,}{flag}")

                        if min_price < threshold:
                            now_ts = time.time()
                            if now_ts - last_alert_time >= cooldown:
                                if sckey and "SendKey" not in sckey:
                                    title = f"{item}低价提醒：{min_price:,}"
                                    desp  = (f"当前最低价 **{min_price:,}** < 阈值 {threshold:,}\n\n"
                                             f"时间：{now_str}")
                                    if send_wechat_alert(sckey, title, desp):
                                        last_alert_time = now_ts
                                else:
                                    print(f"[{now_str}] [警告] serverchan_sckey 未配置，跳过推送")
                            else:
                                remaining_cd = int(cooldown - (time.time() - last_alert_time))
                                print(f"[{now_str}] 推送冷却中，还剩 {remaining_cd}s")
                    else:
                        print(f"[{now_str}] 未识别到价格（货架可能为空）")

                    navigate_back(hwnd, config, parent)
                else:
                    print(f"[{now_str}] 导航失败，跳过本轮")

            elapsed = time.time() - cycle_start
            sleep_time = max(0.0, interval - elapsed)
            if sleep_time > 0:
                time.sleep(sleep_time)

    except KeyboardInterrupt:
        print("\n[监控] 已手动停止。")


def main() -> None:
    config = load_config()
    monitor_loop(config)


if __name__ == "__main__":
    main()
