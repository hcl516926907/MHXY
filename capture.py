import win32gui
import win32con
import ctypes
import ctypes.wintypes
import numpy as np
import mss
import cv2
import os
from typing import Optional, Tuple


def find_game_window() -> Optional[Tuple[int, int, int, int]]:
    """
    Find the game window and return its CLIENT AREA in screen coordinates
    as (x, y, width, height).

    Using the client rect (not GetWindowRect) avoids including Win11's
    transparent window shadow, which would shift all ROI coordinates.
    """
    found = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "梦幻西游" in title:
                found.append(hwnd)

    win32gui.EnumWindows(callback, None)

    if not found:
        return None

    hwnd = found[0]

    # GetClientRect returns (0, 0, client_w, client_h) — relative to window
    _, _, client_w, client_h = win32gui.GetClientRect(hwnd)

    # ClientToScreen converts (0,0) of the client area to absolute screen coords
    pt = ctypes.wintypes.POINT(0, 0)
    ctypes.windll.user32.ClientToScreen(hwnd, ctypes.byref(pt))

    return (pt.x, pt.y, client_w, client_h)


def capture_window(rect: Tuple[int, int, int, int]) -> np.ndarray:
    """Capture the given screen rect using mss. Returns BGR image."""
    x, y, w, h = rect
    with mss.mss() as sct:
        monitor = {"left": x, "top": y, "width": w, "height": h}
        screenshot = sct.grab(monitor)
        img = np.array(screenshot)
        return cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)


def crop_item_grid(img: np.ndarray, roi_ratios: dict) -> np.ndarray:
    """Crop the item grid from the full window image using ratio-based ROI."""
    h, w = img.shape[:2]
    left = int(w * roi_ratios["left"])
    top = int(h * roi_ratios["top"])
    right = int(w * roi_ratios["right"])
    bottom = int(h * roi_ratios["bottom"])
    return img[top:bottom, left:right]


def save_debug_image(img: np.ndarray, label: str = "debug"):
    """Save an image to the debug folder for ROI verification.
    Uses imencode+Python file I/O to avoid cv2.imwrite's lack of Unicode path support."""
    debug_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug")
    os.makedirs(debug_dir, exist_ok=True)
    from datetime import datetime
    ts = datetime.now().strftime("%H%M%S")
    path = os.path.join(debug_dir, f"{label}_{ts}.png")
    ok, buf = cv2.imencode(".png", img)
    if ok:
        with open(path, "wb") as f:
            f.write(buf.tobytes())
    return path
