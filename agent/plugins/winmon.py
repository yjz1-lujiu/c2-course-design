"""
示例插件：窗口信息采集模块 (拓展功能)
通过 Windows API 采集当前窗口标题和用户操作记录
"""
import ctypes
import ctypes.wintypes
import time
import json

user32 = ctypes.windll.user32


def get_foreground_window_title() -> str:
    """获取当前前台窗口标题"""
    try:
        hwnd = user32.GetForegroundWindow()
        length = user32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return ""
        buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buf, length + 1)
        return buf.value
    except Exception:
        return ""


def get_window_list() -> list:
    """枚举所有可见窗口"""
    windows = []

    def enum_callback(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            length = user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buf, length + 1)
                title = buf.value
                if title and len(title) > 1:
                    # 获取窗口进程 ID
                    pid = ctypes.wintypes.DWORD()
                    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                    windows.append({
                        "title": title,
                        "pid": pid.value,
                        "hwnd": hwnd,
                    })
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
    return windows


def monitor_window_changes(duration: int = 60, interval: float = 1.0) -> list:
    """
    监控窗口切换

    Args:
        duration: 监控时长（秒）
        interval: 采样间隔（秒）

    Returns:
        窗口切换记录列表
    """
    records = []
    last_title = ""
    start_time = time.time()

    while time.time() - start_time < duration:
        title = get_foreground_window_title()
        if title and title != last_title:
            records.append({
                "timestamp": time.time(),
                "title": title,
                "duration": 0.0,
            })
            if records and len(records) > 1:
                records[-2]["duration"] = round(
                    time.time() - records[-2]["timestamp"], 2
                )
            last_title = title
        time.sleep(interval)

    # 最后一条的持续时间
    if records:
        records[-1]["duration"] = round(
            time.time() - records[-1]["timestamp"], 2
        )

    return records


def run(duration: int = 60) -> dict:
    """插件入口"""
    windows = get_window_list()
    if duration > 0:
        records = monitor_window_changes(duration)
    else:
        records = []

    return {
        "current_windows": windows,
        "window_switches": records,
    }


if __name__ == "__main__":
    print(json.dumps(run(duration=10), indent=2, ensure_ascii=False))
