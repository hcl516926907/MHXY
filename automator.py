"""
Win32 输入模拟模块

点击（click）：使用 PostMessage 直接向窗口消息队列注入 WM_LBUTTONDOWN/UP，
               绕开硬件输入栈，规避反外挂的 SendInput hook。

拖拽（drag） ：使用 mouse_event() 走硬件输入管道，确保手游移植引擎能响应
               按下→移动→抬起序列（WM_MOUSEWHEEL 对此类游戏通常无效）。
"""

import ctypes
import ctypes.wintypes
import time
from typing import Optional, Tuple

import win32api
import win32con
import win32gui


def get_game_hwnd() -> Optional[int]:
    """返回梦幻西游游戏窗口句柄，未找到则返回 None。"""
    found = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and "梦幻西游" in win32gui.GetWindowText(hwnd):
            found.append(hwnd)

    win32gui.EnumWindows(callback, None)
    return found[0] if found else None


def client_to_screen(hwnd: int, cx: int, cy: int) -> Tuple[int, int]:
    """将客户区坐标转换为屏幕绝对坐标。"""
    pt = ctypes.wintypes.POINT(cx, cy)
    ctypes.windll.user32.ClientToScreen(hwnd, ctypes.byref(pt))
    return pt.x, pt.y


def click(hwnd: int, cx: int, cy: int):
    """
    在客户区坐标 (cx, cy) 处模拟左键单击。
    将客户区坐标转为屏幕坐标后，使用 mouse_event() 经硬件输入管道发送，
    避免 PostMessage 对高完整性级别窗口的"拒绝访问"问题。
    """
    sx, sy = client_to_screen(hwnd, cx, cy)
    win32api.SetCursorPos((sx, sy))
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)


def drag(x_start: int, y_start: int, x_end: int, y_end: int,
         steps: int = 20, duration: float = 0.4, inertia: bool = False):
    """
    从 (x_start, y_start) 拖拽到 (x_end, y_end)，使用屏幕绝对坐标。
    使用 mouse_event() 经硬件输入管道发送，手游移植引擎可识别。

    inertia=False（默认）：到达终点后静止 0.3 秒再抬起，引擎判定为慢速拖拽，无惯性。
    inertia=True         ：到达终点立即抬起，引擎判定为快速抛掷，触发惯性继续滑动。
    """
    win32api.SetCursorPos((x_start, y_start))
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.05)
    for i in range(1, steps + 1):
        x = x_start + (x_end - x_start) * i // steps
        y = y_start + (y_end - y_start) * i // steps
        win32api.SetCursorPos((x, y))
        time.sleep(duration / steps)
    if not inertia:
        # 静止停顿，消除惯性
        time.sleep(0.3)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
