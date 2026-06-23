"""
沙箱/虚拟机/调试器检测模块
通过 ctypes 调用 Windows API 实现环境风险评估
"""
import ctypes
import ctypes.wintypes
import os
import time
import string

# Windows API 句柄
kernel32 = ctypes.windll.kernel32
user32 = ctypes.windll.user32
advapi32 = ctypes.windll.advapi32

# ── 常量定义 ──────────────────────────────────────────────

CURSORINFO_STRUCT_SIZE = ctypes.sizeof(ctypes.wintypes.DWORD) * 2 + ctypes.sizeof(ctypes.wintypes.POINT) + ctypes.sizeof(ctypes.wintypes.HANDLE)
GENERIC_READ = 0x80000000
KEY_READ = 0x20019
REG_SZ = 1
HKEY_LOCAL_MACHINE = 0x80000002


class CURSORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.wintypes.DWORD),
        ("flags", ctypes.wintypes.DWORD),
        ("ptScreenPos", ctypes.wintypes.POINT),
        ("hCursor", ctypes.wintypes.HANDLE),
    ]


# ── 虚拟机特征库 ──────────────────────────────────────────

VM_REGISTRY_KEYS = [
    (HKEY_LOCAL_MACHINE, r"SOFTWARE\VMware, Inc.\VMware Tools"),
    (HKEY_LOCAL_MACHINE, r"SOFTWARE\Oracle\VirtualBox Guest Additions"),
    (HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Services\vmci"),
    (HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Services\VBoxGuest"),
]

VM_PROCESSES = [
    "vmtoolsd.exe", "vmwaretray.exe", "vmwareuser.exe",
    "VBoxService.exe", "VBoxTray.exe",
    "vmacthlp.exe", "vmware.exe",
]

VM_MAC_PREFIXES = [
    b"\x00\x05\x69",  # VMware
    b"\x00\x0c\x29",  # VMware
    b"\x00\x50\x56",  # VMware
    b"\x08\x00\x27",  # VirtualBox
]

DEBUGGER_PROCESSES = [
    "x64dbg.exe", "x32dbg.exe", "ollydbg.exe", "ida.exe", "ida64.exe",
    "windbg.exe", "procmon.exe", "wireshark.exe", "fiddler.exe",
    "processhacker.exe", "dumpcap.exe",
]


class SandboxDetector:
    """沙箱/VM/调试器检测器"""

    def __init__(self):
        self.checks = []
        self.risk_score = 0
        self.max_score = 0

    def _add_check(self, name: str, detected: bool, weight: int = 1):
        self.checks.append({"name": name, "detected": detected, "weight": weight})
        self.max_score += weight
        if detected:
            self.risk_score += weight

    # ── 调试器检测 ────────────────────────────────────────

    def check_debugger_present(self) -> bool:
        """IsDebuggerPresent API 检测"""
        try:
            result = kernel32.IsDebuggerPresent()
            return bool(result)
        except Exception:
            return False

    def check_remote_debugger(self) -> bool:
        """CheckRemoteDebuggerPresent 检测"""
        try:
            is_debugged = ctypes.wintypes.BOOL(False)
            handle = kernel32.GetCurrentProcess()
            kernel32.CheckRemoteDebuggerPresent(handle, ctypes.byref(is_debugged))
            return bool(is_debugged)
        except Exception:
            return False

    def check_debugger_processes(self) -> bool:
        """检测调试器进程是否存在"""
        try:
            import subprocess
            output = subprocess.check_output(
                "tasklist /FO CSV /NH", shell=True, timeout=5
            ).decode("utf-8", errors="ignore").lower()
            for proc in DEBUGGER_PROCESSES:
                if proc.lower() in output:
                    return True
        except Exception:
            pass
        return False

    # ── 虚拟机检测 ────────────────────────────────────────

    def check_vm_registry(self) -> bool:
        """通过注册表键值检测 VM"""
        for hive, path in VM_REGISTRY_KEYS:
            try:
                hkey = ctypes.wintypes.HKEY()
                ret = advapi32.RegOpenKeyExW(
                    hive, path, 0, KEY_READ, ctypes.byref(hkey)
                )
                if ret == 0:
                    advapi32.RegCloseKey(hkey)
                    return True
            except Exception:
                continue
        return False

    def check_vm_processes(self) -> bool:
        """检测 VM 相关进程"""
        try:
            import subprocess
            output = subprocess.check_output(
                "tasklist /FO CSV /NH", shell=True, timeout=5
            ).decode("utf-8", errors="ignore").lower()
            for proc in VM_PROCESSES:
                if proc.lower() in output:
                    return True
        except Exception:
            pass
        return False

    def check_vm_mac(self) -> bool:
        """通过 MAC 地址前缀检测 VM"""
        try:
            import subprocess
            output = subprocess.check_output(
                "getmac /FO CSV /NH", shell=True, timeout=5
            ).decode("utf-8", errors="ignore")
            for line in output.strip().split("\n"):
                mac_str = line.split(",")[0].strip('"').replace("-", ":")
                parts = mac_str.split(":")
                if len(parts) >= 3:
                    prefix = bytes(int(p, 16) for p in parts[:3])
                    if prefix in VM_MAC_PREFIXES:
                        return True
        except Exception:
            pass
        return False

    def check_vm_disk_size(self) -> bool:
        """磁盘容量检测（沙箱通常磁盘较小）"""
        try:
            total = 0
            for letter in string.ascii_uppercase:
                drive = f"{letter}:\\"
                if os.path.exists(drive):
                    import shutil
                    _, _, free = shutil.disk_usage(drive)
                    total += free
            # 小于 60GB 可疑
            return total < 60 * 1024 * 1024 * 1024
        except Exception:
            return False

    # ── 沙箱行为检测 ──────────────────────────────────────

    def check_uptime(self) -> bool:
        """运行时间检测（沙箱运行时间通常很短）"""
        try:
            tick = kernel32.GetTickCount64()
            # 少于 10 分钟可疑
            return tick < 10 * 60 * 1000
        except Exception:
            return False

    def check_mouse_movement(self) -> bool:
        """鼠标移动检测（沙箱通常无鼠标操作）"""
        try:
            positions = []
            for _ in range(3):
                ci = CURSORINFO()
                ci.cbSize = ctypes.sizeof(CURSORINFO)
                user32.GetCursorInfo(ctypes.byref(ci))
                positions.append((ci.ptScreenPos.x, ci.ptScreenPos.y))
                time.sleep(0.3)
            # 三次采样位置完全相同 -> 无人操作
            return len(set(positions)) == 1
        except Exception:
            return False

    def check_keyboard_activity(self) -> bool:
        """键盘输入频率检测"""
        try:
            key_count = 0
            for vk in range(8, 256):
                if user32.GetAsyncKeyState(vk) & 0x8000:
                    key_count += 1
            return key_count == 0
        except Exception:
            return False

    def check_desktop_processes(self) -> bool:
        """桌面进程活跃度检测"""
        try:
            import subprocess
            output = subprocess.check_output(
                "tasklist /FO CSV /NH", shell=True, timeout=5
            ).decode("utf-8", errors="ignore").lower()
            desktop_procs = ["explorer.exe"]
            for proc in desktop_procs:
                if proc not in output:
                    return True
            # 检查窗口数量
            hwnd_count = user32.GetForegroundWindow()
            return hwnd_count == 0
        except Exception:
            return False

    # ── 综合评估 ──────────────────────────────────────────

    def run_all_checks(self) -> dict:
        """运行所有检测，返回风险评估结果"""
        self.checks = []
        self.risk_score = 0
        self.max_score = 0

        # 调试器检测 (权重高)
        self._add_check("is_debugger_present", self.check_debugger_present(), 3)
        self._add_check("remote_debugger", self.check_remote_debugger(), 3)
        self._add_check("debugger_processes", self.check_debugger_processes(), 2)

        # VM 检测
        self._add_check("vm_registry", self.check_vm_registry(), 2)
        self._add_check("vm_processes", self.check_vm_processes(), 2)
        self._add_check("vm_mac", self.check_vm_mac(), 2)

        # 沙箱行为
        self._add_check("short_uptime", self.check_uptime(), 1)
        self._add_check("no_mouse_movement", self.check_mouse_movement(), 1)
        self._add_check("no_keyboard", self.check_keyboard_activity(), 1)
        self._add_check("low_disk", self.check_vm_disk_size(), 1)
        self._add_check("no_desktop", self.check_desktop_processes(), 1)

        # 计算风险等级
        ratio = self.risk_score / max(self.max_score, 1)
        if ratio >= 0.5:
            level = "sandbox"
        elif ratio >= 0.25:
            level = "suspicious"
        else:
            level = "safe"

        return {
            "risk_score": self.risk_score,
            "max_score": self.max_score,
            "risk_ratio": round(ratio, 2),
            "level": level,
            "checks": self.checks,
        }


def detect() -> dict:
    """便捷入口：运行沙箱检测"""
    detector = SandboxDetector()
    return detector.run_all_checks()


if __name__ == "__main__":
    import json
    result = detect()
    print(json.dumps(result, indent=2, ensure_ascii=False))
