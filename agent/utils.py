"""
Agent 工具函数
"""
import os
import sys
import hashlib
import platform
import logging

logger = logging.getLogger("agent.utils")


def is_admin() -> bool:
    """检测是否以管理员权限运行"""
    try:
        import ctypes
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def get_temp_dir() -> str:
    """获取临时目录路径"""
    return os.environ.get("TEMP", os.environ.get("TMP", "."))


def secure_delete(filepath: str):
    """
    安全删除文件（覆写后删除）
    注意：SSD 上由于 wear leveling，此方法效果有限
    """
    try:
        if not os.path.exists(filepath):
            return
        size = os.path.getsize(filepath)
        with open(filepath, "wb") as f:
            # 全零覆写
            f.write(b"\x00" * size)
            f.flush()
            os.fsync(f.fileno())
        os.remove(filepath)
        logger.info(f"Securely deleted: {filepath}")
    except Exception as e:
        logger.error(f"Secure delete failed for {filepath}: {e}")


def file_hash(filepath: str, algo: str = "sha256") -> str:
    """计算文件哈希"""
    h = hashlib.new(algo)
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def get_system_info() -> dict:
    """获取系统基础信息"""
    return {
        "hostname": platform.node(),
        "os": platform.platform(),
        "version": platform.version(),
        "arch": platform.machine(),
        "processor": platform.processor(),
        "python": platform.python_version(),
        "pid": os.getpid(),
        "cwd": os.getcwd(),
        "user": os.environ.get("USERNAME", os.environ.get("USER", "unknown")),
    }


def hide_console_window():
    """隐藏控制台窗口（仅 Windows）"""
    try:
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE
    except Exception:
        pass
