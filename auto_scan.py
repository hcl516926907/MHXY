"""
自动化商城扫描

流程：
  1. 截图 + OCR 识别当前分类选择界面，找出可见商品名称及坐标
  2. 逐一点击 target_categories 中的目标 → 等待加载 → scan() → 点击父分类返回
  3. 若当前视图没有剩余目标 → 拖动列表 → 重复
  4. 直到所有目标扫完或超过最大拖动次数

target_categories 格式（config.json）：
  {"高级兽诀": ["高级必杀", "高级连击", "高级强力"],
   "低级兽诀": ["低级必杀", "低级连击"]}
"""

import time
from typing import Callable, Dict, List, Tuple

import cv2

from automator import client_to_screen, click, drag, get_game_hwnd
from capture import capture_window, find_game_window, save_debug_image
from ocr import recognize_with_positions
from ocr_corrections import normalize_ocr_text
from storage import save_na_result


def _norm(text: str) -> str:
    """归一化商品名：OCR 字符修正 + 去掉中文点（·/・）和空格，用于模糊匹配。"""
    text = normalize_ocr_text(text)
    return text.replace("·", "").replace("・", "").replace(" ", "").strip()


def _find_categories_on_screen(
    full_img,
    config: dict,
    roi_key: str = "category_roi_ratios",
) -> Dict[str, Tuple[int, int]]:
    """
    对指定 ROI 区域做 OCR，返回 {文本: (client_cx, client_cy)}。
    坐标为窗口客户区像素坐标，可直接用于 click()。

    roi_key : config 中使用的 ROI 键名，默认 "category_roi_ratios"（右侧子商品面板）；
              传入 "back_roi_ratios" 则扫描左侧父分类导航面板。
    """
    roi = config[roi_key]
    h, w = full_img.shape[:2]
    crop_x      = int(w * roi["left"])
    crop_y      = int(h * roi["top"])
    crop_right  = int(w * roi["right"])
    crop_bottom = int(h * roi["bottom"])
    crop = full_img[crop_y:crop_bottom, crop_x:crop_right]

    results = recognize_with_positions(
        crop,
        use_gpu=config.get("use_gpu", False),
        confidence_threshold=config.get("ocr_confidence_threshold", 0.7),
    )

    categories: Dict[str, Tuple[int, int]] = {}
    for text, _conf, cx_in_crop, cy_in_crop in results:
        text = text.strip()
        if len(text) >= 1:
            # 将裁剪图坐标还原为客户区坐标；同时以归一化名（去点）为 key 存一份
            coord = (crop_x + cx_in_crop, crop_y + cy_in_crop)
            categories[text] = coord
            norm_key = _norm(text)
            if norm_key != text:          # 避免重复覆盖同名条目
                categories[norm_key] = coord

    return categories


def auto_scan(
    config: dict,
    server_name: str,
    server_open_dt,
    scan_fn: Callable,
):
    """
    自动扫描 config["target_categories"] 中所有商品的价格。

    target_categories 必须为字典格式：{"父分类名": ["子商品1", "子商品2", ...]}
    扫完每个子商品后，程序会通过 OCR 找到父分类名并点击（返回分类选择界面），
    无需手动设置返回按钮坐标。

    参数
    ----
    scan_fn : main.scan 函数的引用，签名为 scan(config, server_name, server_open_dt)
    """
    targets = config.get("target_categories", {})

    if isinstance(targets, list):
        print("[自动扫描] target_categories 格式有误，请改为字典格式：")
        print('  {"高级兽诀": ["高级必杀", "高级连击"], "低级兽诀": ["低级必杀"]}')
        return

    if not targets:
        print("[自动扫描] config.json 中 target_categories 为空，已跳过。")
        print('  请设置：{"父分类": ["子商品1", "子商品2"]}')
        return

    max_drags: int   = config.get("auto_scan_max_drags", 5)
    scan_wait: float = config.get("scan_wait", 1.0)
    back_wait: float = config.get("back_wait", 0.6)
    drag_wait: float = config.get("drag_wait", 1.0)

    total = sum(len(v) for v in targets.values())
    print(f"\n[自动扫描] 开始，共 {len(targets)} 个父分类，{total} 个子商品")
    for parent, subs in targets.items():
        print(f"  {parent}：{subs}")

    hwnd = get_game_hwnd()
    if hwnd is None:
        print("[自动扫描] 未找到梦幻西游窗口，已停止。")
        return

    for parent, sub_list in targets.items():
        remaining = list(sub_list)
        drags_done = 0
        print(f"\n  ── 父分类「{parent}」，共 {len(sub_list)} 个子商品 ──")

        # 先点击父分类，进入该分类的子商品列表界面
        _click_parent_via_ocr(hwnd, parent, config)
        time.sleep(config.get("back_wait", 0.6))

        # 向下拖拽（带惯性）回到列表顶部，使用独立的大距离配置
        rect = find_game_window()
        if rect is not None:
            _, _, w, h = rect
            drag_cfg   = config.get("drag_ratio", {"x": 0.62, "y_start": 0.65, "y_end": 0.30})
            top_cfg    = config.get("scroll_to_top_ratio",
                                    {"y_start": 0.75, "y_end": 0.10})  # 默认跨 65% 窗口高
            cx_drag    = int(w * drag_cfg["x"])
            sx, sy_from = client_to_screen(hwnd, cx_drag, int(h * top_cfg["y_start"]))
            _,  sy_to   = client_to_screen(hwnd, cx_drag, int(h * top_cfg["y_end"]))
            scroll_top_times = config.get("scroll_to_top_times", 2)
            print(f"    ↓ 回到顶部（向下拖拽 {scroll_top_times} 次）")
            for _ in range(scroll_top_times):
                drag(sx, sy_from, sx, sy_to)
                time.sleep(drag_wait)

        items_per_page: int = config.get("items_per_page", 6)
        scanned_this_page   = 0   # 本页已扫件数，达到 items_per_page 时触发拖拽

        while remaining:
            rect = find_game_window()
            if rect is None:
                print("[自动扫描] 游戏窗口已关闭，已停止。")
                return
            _, _, w, h = rect

            full_img = capture_window(rect)
            visible = _find_categories_on_screen(full_img, config)
            print(f"    [OCR] 识别到 {len(visible)} 个分类：{list(visible.keys())}")

            # 用归一化名匹配（忽略·），找出本轮可点击的目标
            to_click = []
            for t in remaining:
                key = _norm(t) if _norm(t) in visible else t
                if key in visible:
                    to_click.append((t, visible[key]))

            if to_click:
                for target, (cx, cy) in to_click:
                    print(f"    → 点击「{target}」（客户区 {cx}, {cy}）")
                    click(hwnd, cx, cy)
                    time.sleep(scan_wait)

                    found = scan_fn(config, server_name, server_open_dt, category=parent)
                    if found is False:
                        na_path = save_na_result(target, config["output_dir"], server_name, server_open_dt, category=parent)
                        print(f"      货架为空，已记录 NA：{na_path}")

                    _click_back_via_ocr(hwnd, parent, config)
                    time.sleep(back_wait)

                    remaining.remove(target)
                    scanned_this_page += 1

                    # 每扫满一页就拖动一次，避免第 7/8 件因半截而识别失败
                    if scanned_this_page >= items_per_page and remaining:
                        scanned_this_page = 0
                        drags_done += 1
                        _do_drag(hwnd, config, w, h, full_img, drag_wait)
                        break   # 重新进入循环，用新截图重新 OCR

            else:
                # 当前视图没有剩余目标，尝试拖动列表
                if drags_done >= max_drags:
                    print(f"    [警告] 已拖动 {drags_done} 次，仍未找到：{remaining}")
                    print("      请检查：1. 名称是否与游戏内完全一致")
                    print("              2. drag_ratio / category_roi_ratios 是否需要校准")
                    break

                drags_done += 1
                _do_drag(hwnd, config, w, h, full_img, drag_wait)

    print(f"\n[自动扫描] 完成！共扫描 {total} 个子商品。")


def _do_drag(hwnd: int, config: dict, w: int, h: int, full_img, drag_wait: float):
    """执行一次子商品列表的上滑拖拽，并保存 debug 图。"""
    import numpy as np
    drag_cfg = config.get("drag_ratio", {"x": 0.62, "y_start": 0.65, "y_end": 0.30})
    cx_drag  = int(w * drag_cfg["x"])
    cy_start = int(h * drag_cfg["y_start"])
    cy_end   = int(h * drag_cfg["y_end"])
    sx, sy_start = client_to_screen(hwnd, cx_drag, cy_start)
    _,  sy_end   = client_to_screen(hwnd, cx_drag, cy_end)

    drag_debug = full_img.copy()
    orange = (0, 165, 255)
    cv2.circle(drag_debug, (cx_drag, cy_start), 12, orange, -1)
    tri_pts = [
        (cx_drag,      cy_end - 14),
        (cx_drag - 12, cy_end + 10),
        (cx_drag + 12, cy_end + 10),
    ]
    cv2.fillPoly(drag_debug, [np.array(tri_pts, dtype=np.int32)], orange)
    save_debug_image(drag_debug, "full_window_drag")

    print(f"    ↕ 拖动列表（屏幕：({sx}, {sy_start}) → ({sx}, {sy_end})）")
    drag(sx, sy_start, sx, sy_end)
    time.sleep(drag_wait)


def _drag_back_roi(hwnd: int, config: dict, w: int, h: int):
    """在 back_roi 区域内向下拖拽（使左侧导航列表向下滚动，露出更多父分类）。"""
    import numpy as np
    roi = config.get("back_roi_ratios", {"left": 0.12, "top": 0.18, "right": 0.22, "bottom": 0.55})
    # 拖拽 x：ROI 水平中心
    cx_drag = int(w * (roi["left"] + roi["right"]) / 2)
    # 从 ROI 底部 80% 处拖到 ROI 顶部 20% 处（向上拖 = 列表向下滚动）
    cy_from = int(h * (roi["top"] + (roi["bottom"] - roi["top"]) * 0.8))
    cy_to   = int(h * (roi["top"] + (roi["bottom"] - roi["top"]) * 0.2))
    sx, sy_from = client_to_screen(hwnd, cx_drag, cy_from)
    _,  sy_to   = client_to_screen(hwnd, cx_drag, cy_to)

    # debug：在 full_window 截图上用橙色标注 back_roi 拖拽起止点
    rect = find_game_window()
    if rect is not None:
        full_img = capture_window(rect)
        dbg = full_img.copy()
        orange = (0, 165, 255)
        cv2.circle(dbg, (cx_drag, cy_from), 10, orange, -1)
        tri_pts = [
            (cx_drag,      cy_to - 12),
            (cx_drag - 10, cy_to + 8),
            (cx_drag + 10, cy_to + 8),
        ]
        cv2.fillPoly(dbg, [np.array(tri_pts, dtype=np.int32)], orange)
        save_debug_image(dbg, "full_window_back_drag")

    drag(sx, sy_from, sx, sy_to)


def _click_parent_via_ocr(hwnd: int, parent: str, config: dict):
    """
    在分类选择界面，OCR 左侧导航面板找到父分类名并点击，进入子商品列表。
    若未找到，向下滚动 back_roi 列表后重试，最多重试 auto_scan_max_drags 次。
    """
    max_drags: int   = config.get("auto_scan_max_drags", 5)
    drag_wait: float = config.get("drag_wait", 1.0)

    for attempt in range(max_drags + 1):
        rect = find_game_window()
        if rect is None:
            return
        _, _, w, h = rect

        full_img = capture_window(rect)
        visible = _find_categories_on_screen(full_img, config, roi_key="back_roi_ratios")

        if parent in visible:
            cx, cy = visible[parent]
            print(f"    → 点击父分类「{parent}」（客户区 {cx}, {cy}）")
            click(hwnd, cx, cy)
            return

        if attempt < max_drags:
            print(f"    → OCR 未找到父分类「{parent}」，向下滚动 back_roi 后重试（{attempt + 1}/{max_drags}）")
            _drag_back_roi(hwnd, config, w, h)
            time.sleep(drag_wait)

    # 所有重试耗尽，fallback 固定坐标
    rect = find_game_window()
    if rect is None:
        return
    _, _, w, h = rect
    back = config.get("back_button_ratio", {"x": 0.18, "y": 0.49})
    cx = int(w * back["x"])
    cy = int(h * back["y"])
    print(f"    → 多次重试仍未找到「{parent}」，使用备用坐标（客户区 {cx}, {cy}）")
    click(hwnd, cx, cy)


def _click_back_via_ocr(hwnd: int, parent: str, config: dict):
    """
    截图并 OCR 当前界面，找到父分类名称后点击（返回分类选择界面）。
    若 OCR 未能识别父分类名，则 fallback 到 back_button_ratio 固定坐标。
    """
    rect = find_game_window()
    if rect is None:
        return
    _, _, w, h = rect

    full_img = capture_window(rect)

    # ── 调试图：在全窗口截图上用蓝框标出 back_roi 区域，并单独保存裁剪图 ──
    roi = config.get("back_roi_ratios", {"left": 0.12, "top": 0.18, "right": 0.22, "bottom": 0.55})
    img_h, img_w = full_img.shape[:2]
    bx1, by1 = int(img_w * roi["left"]),  int(img_h * roi["top"])
    bx2, by2 = int(img_w * roi["right"]), int(img_h * roi["bottom"])
    annotated = full_img.copy()
    cv2.rectangle(annotated, (bx1, by1), (bx2, by2), (255, 0, 0), 3)   # 蓝框
    save_debug_image(annotated, "full_window_back_roi")
    save_debug_image(full_img[by1:by2, bx1:bx2], "back_roi")
    print(f"    [调试] back_roi 截图已保存（像素 x={bx1}~{bx2}, y={by1}~{by2}）")
    # ─────────────────────────────────────────────────────────────────────────

    # 用左侧导航面板的 ROI 扫父分类（back_roi_ratios），而非右侧子商品面板
    visible = _find_categories_on_screen(full_img, config, roi_key="back_roi_ratios")

    if parent in visible:
        cx, cy = visible[parent]
        print(f"    ← 返回「{parent}」（OCR 识别，客户区 {cx}, {cy}）")
        click(hwnd, cx, cy)
    else:
        # fallback：使用固定比例坐标
        back = config.get("back_button_ratio", {"x": 0.18, "y": 0.49})
        cx = int(w * back["x"])
        cy = int(h * back["y"])
        print(f"    ← OCR 未找到「{parent}」，使用备用坐标（客户区 {cx}, {cy}）")
        click(hwnd, cx, cy)
