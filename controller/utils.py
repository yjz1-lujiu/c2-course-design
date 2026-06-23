"""
控制端工具函数
"""
import os
import json
import time
import hashlib
import logging

logger = logging.getLogger("controller.utils")


def format_size(size: int) -> str:
    """格式化文件大小"""
    for unit in ["B", "KB", "MB", "GB"]:
        if size < 1024:
            return f"{size:.1f}{unit}"
        size /= 1024
    return f"{size:.1f}TB"


def format_duration(seconds: float) -> str:
    """格式化时间间隔"""
    if seconds < 60:
        return f"{seconds:.0f}s"
    elif seconds < 3600:
        return f"{seconds / 60:.0f}m {seconds % 60:.0f}s"
    else:
        h = seconds // 3600
        m = (seconds % 3600) // 60
        return f"{h:.0f}h {m:.0f}m"


def save_session_log(session_id: str, data: dict, log_dir: str = "logs"):
    """保存会话日志"""
    os.makedirs(log_dir, exist_ok=True)
    filename = f"{log_dir}/session_{session_id[:8]}_{int(time.time())}.json"
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    logger.info(f"Session log saved: {filename}")
    return filename


def print_table(headers: list, rows: list, col_width: int = 20):
    """打印格式化表格"""
    header_str = "".join(h.ljust(col_width) for h in headers)
    print(f"\n{header_str}")
    print("-" * (col_width * len(headers)))
    for row in rows:
        row_str = "".join(str(c).ljust(col_width) for c in row)
        print(row_str)
    print()
