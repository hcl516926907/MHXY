"""
梦幻西游手游 - 商城价格扫描工具
用法：以管理员权限运行，在游戏摆摊界面按 F9 扫描当前页价格。
"""

# ── 必须在所有其他 import 之前设置 ─────────────────────────────────────────
import os
import sys
import types
import ctypes

# 0. 屏蔽 modelscope：paddlex 会无条件 import modelscope，modelscope 又会 import torch，
#    而本机 torch 的 shm.dll 依赖缺失会导致 OSError [WinError 127]。
#    推断模型以本地 PaddlePaddle 运行，不需要 ModelScope hub，注入空 stub 即可绕过。
if "modelscope" not in sys.modules:
    _ms_stub = types.ModuleType("modelscope")
    sys.modules["modelscope"] = _ms_stub

# 1. DPI 感知：外接屏场景下确保 GetWindowRect 返回物理像素坐标
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass

# 2. 禁用 oneDNN/MKL-DNN：PaddlePaddle 3.x CPU 模式下 oneDNN 存在 PIR 指令兼容 bug，
#    ConvertPirAttribute2RuntimeAttribute 会抛 NotImplementedError，禁用后回退标准实现
os.environ["FLAGS_use_mkldnn"] = "0"
os.environ["FLAGS_enable_pir_api"] = "0"   # 禁用 PIR 执行器，避免 oneDNN 指令兼容 bug
# ────────────────────────────────────────────────────────────────────────────

import cv2
import keyboard

from config import load_config
from capture import find_game_window, capture_window, crop_item_grid, save_debug_image, set_debug_enabled
from ocr import recognize
from parser import parse_page
from storage import save_result
from servers import select_server
from auto_scan import auto_scan


def scan(config: dict, server_name: str, server_open_dt, category: str = ""):
    print("\n[扫描中...]")

    # 1. 找游戏窗口
    rect = find_game_window()
    if rect is None:
        print("[错误] 未找到梦幻西游窗口，请确认游戏正在运行。")
        return
    x, y, w, h = rect
    print(f"  游戏窗口：({x}, {y})  {w}x{h}")

    # 2. 截图
    full_img = capture_window(rect)

    # 3. 计算 ROI 像素坐标并打印（用于确认配置是否正确加载）
    roi = config["roi_ratios"]
    img_h, img_w = full_img.shape[:2]
    roi_left   = int(img_w * roi["left"])
    roi_top    = int(img_h * roi["top"])
    roi_right  = int(img_w * roi["right"])
    roi_bottom = int(img_h * roi["bottom"])
    print(f"  截图尺寸：{img_w}x{img_h}")
    print(f"  ROI 比例：left={roi['left']} top={roi['top']} right={roi['right']} bottom={roi['bottom']}")
    print(f"  ROI 像素：x={roi_left}~{roi_right}, y={roi_top}~{roi_bottom}  ({roi_right-roi_left}x{roi_bottom-roi_top})")

    # 在 full_window 上叠加 ROI 红框后保存（方便目测偏差位置）
    annotated = full_img.copy()
    cv2.rectangle(annotated, (roi_left, roi_top), (roi_right, roi_bottom), (0, 0, 255), 3)
    save_debug_image(annotated, "full_window_roi")

    # 4. 裁剪物品网格区域
    grid_img = crop_item_grid(full_img, config["roi_ratios"])

    # 保存 ROI 裁剪截图
    debug_path = save_debug_image(grid_img, "roi")
    print(f"  调试截图：debug/full_window_roi_*.png（红框标注位置）和 debug/roi_*.png（裁剪结果）")

    # 4. OCR 识别
    texts = recognize(
        grid_img,
        use_gpu=config.get("use_gpu", True),
        confidence_threshold=config.get("ocr_confidence_threshold", 0.7),
    )
    print(f"  识别到 {len(texts)} 个文字块：")
    for text, conf in texts:
        print(f"    [{conf:.2f}] {text}")

    # 5. 解析物品名和价格
    item_name, prices = parse_page(texts)

    if not prices:
        print("[警告] 未识别到任何价格（货架为空或 ROI 未覆盖商品列表）。")
        print("  请检查：")
        print("  1. ROI 区域是否覆盖商品列表（查看 debug/ 目录中的截图）")
        print("  2. 如需调整，修改 config.json 中的 roi_ratios 值")
        return False

    avg_price = sum(prices) / len(prices)

    # 6. 打印结果
    print(f"\n  物品名称：{item_name or '未识别'}")
    print(f"  价格列表：{[f'{p:,}' for p in sorted(prices)]}")
    print(f"  平均价格：{avg_price:,.0f}")
    print(f"  商品数量：{len(prices)}")

    # 7. 保存到 CSV
    filepath = save_result(item_name, prices, config["output_dir"], server_name, server_open_dt, category)
    print(f"  已保存：{filepath}")
    return True


def _clear_debug_folder():
    debug_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug")
    if os.path.isdir(debug_dir):
        removed = 0
        for fname in os.listdir(debug_dir):
            if fname.lower().endswith((".png", ".jpg", ".jpeg", ".bmp")):
                os.remove(os.path.join(debug_dir, fname))
                removed += 1
        if removed:
            print(f"[初始化] 已清除 debug/ 中 {removed} 张旧截图。")


def main():
    config = load_config()
    hotkey = config.get("hotkey", "f9")

    _clear_debug_folder()

    print("=" * 50)
    print("  梦幻西游 商城价格扫描工具")
    print("=" * 50)
    auto_scan_hotkey = config.get("auto_scan_hotkey", "f10")
    print(f"  热键：{hotkey.upper()} — 手动扫描当前页价格")
    print(f"  热键：{auto_scan_hotkey.upper()} — 自动扫描所有目标商品")
    print(f"  退出：ESC")
    print(f"  ROI：{config['roi_ratios']}")
    print(f"  GPU：{'是' if config.get('use_gpu') else '否'}")
    print(f"  输出：{config['output_dir']}/")
    print("=" * 50)
    print("\n注意：请以管理员权限运行，否则全局热键可能无效。\n")

    # 首次 OCR 调用会加载模型（约数秒），提前预热
    print("正在加载 OCR 模型...")
    from ocr import get_ocr
    get_ocr(use_gpu=config.get("use_gpu", True))
    print("OCR 模型加载完成，可以开始扫描。\n")

    debug_ans = input("是否生成 debug 调试图片？[y/N] ").strip().lower()
    set_debug_enabled(debug_ans == "y")
    print(f"  调试图片：{'开启' if debug_ans == 'y' else '关闭'}")

    server_name, server_open_dt = select_server()
    print(f"\n按 {hotkey.upper()} 手动扫描当前页，按 {auto_scan_hotkey.upper()} 自动扫描所有目标商品。\n")

    keyboard.add_hotkey(hotkey, scan, args=(config, server_name, server_open_dt))
    def _auto_scan_and_exit():
        auto_scan(config, server_name, server_open_dt, scan)
        print("\n[自动扫描] 已完成，程序退出。")
        os._exit(0)

    keyboard.add_hotkey(auto_scan_hotkey, _auto_scan_and_exit)
    keyboard.wait("esc")
    print("\n程序已退出。")


if __name__ == "__main__":
    main()
