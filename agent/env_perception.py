"""
WMI 系统环境深度感知模块
采集主机信息、识别安全软件、检测白名单应用
"""
import os
import ctypes
import platform
from dataclasses import dataclass, field, asdict
from typing import List, Optional


# ── 安全软件进程名列表 ───────────────────────────────────

SECURITY_SOFTWARE = {
    "火绒": ["HipsTray.exe", "HipsMain.exe", "HipsDaemon.exe", "wsctrl.exe"],
    "360": ["360tray.exe", "360sd.exe", "360safe.exe", "ZhuDongFangYu.exe"],
    "腾讯电脑管家": ["QQPCRTP.exe", "QQPCTray.exe"],
    "Windows Defender": ["MsMpEng.exe", "NisSrv.exe", "MpCmdRun.exe"],
    "卡巴斯基": ["avp.exe", "avpui.exe"],
    "Norton": ["NortonSecurity.exe", "ns.exe"],
    "McAfee": ["mcshield.exe", "mctray.exe"],
    "Avast": ["AvastSvc.exe", "AvastUI.exe"],
}

# ── 白名单/常用软件进程名 ─────────────────────────────────

WHITELIST_SOFTWARE = {
    "Outlook": ["OUTLOOK.EXE", "olk.exe"],
    "Word": ["WINWORD.EXE"],
    "Excel": ["EXCEL.EXE"],
    "PowerPoint": ["POWERPNT.EXE"],
    "微信": ["WeChat.exe", "WeChatApp.exe"],
    "钉钉": ["DingTalk.exe", "DingtalkUp.exe"],
    "QQ": ["QQ.exe"],
    "Chrome": ["chrome.exe"],
    "Firefox": ["firefox.exe"],
    "Edge": ["msedge.exe"],
    "Teams": ["Teams.exe", "ms-teams.exe"],
}


@dataclass
class EnvProfile:
    """系统环境画像"""
    hostname: str = ""
    os_version: str = ""
    os_build: str = ""
    arch: str = ""
    memory_mb: int = 0
    disk_free_gb: float = 0.0
    is_admin: bool = False
    username: str = ""
    domain: str = ""
    # 检测到的安全软件
    security_software: List[str] = field(default_factory=list)
    # 检测到的白名单软件
    whitelist_software: List[str] = field(default_factory=list)
    # 所有运行中的进程名
    running_processes: List[str] = field(default_factory=list)

    def has_outlook(self) -> bool:
        return "Outlook" in self.whitelist_software

    def has_office(self) -> bool:
        return any(s in self.whitelist_software
                   for s in ["Word", "Excel", "PowerPoint"])

    def has_wechat(self) -> bool:
        return "微信" in self.whitelist_software

    def has_defender(self) -> bool:
        return "Windows Defender" in self.security_software

    def to_dict(self) -> dict:
        return asdict(self)


class EnvPerception:
    """WMI 环境感知引擎"""

    def __init__(self):
        self.profile = EnvProfile()
        self._wmi = None
        self._use_subprocess = False

    def _init_wmi(self):
        """初始化 WMI 连接"""
        try:
            import wmi
            self._wmi = wmi.WMI()
            self._use_subprocess = False
        except ImportError:
            self._use_subprocess = True

    def _wmi_query(self, wql: str) -> list:
        """执行 WMI 查询，支持 wmi 库和 subprocess 两种方式"""
        if self._wmi and not self._use_subprocess:
            return self._wmi.query(wql)

        # fallback: subprocess 调用 wmic
        import subprocess
        try:
            # 简化 WQL 到 wmic 命令
            class_name = wql.split("FROM ")[-1].split(" ")[0].strip() if "FROM" in wql else wql
            output = subprocess.check_output(
                f"wmic {class_name} get * /FORMAT:LIST",
                shell=True, timeout=10, stderr=subprocess.DEVNULL
            ).decode("utf-8", errors="ignore")
            # 解析 LIST 格式
            results = []
            current = {}
            for line in output.split("\n"):
                line = line.strip()
                if "=" in line:
                    key, _, val = line.partition("=")
                    current[key.strip()] = val.strip()
                elif current:
                    results.append(current)
                    current = {}
            if current:
                results.append(current)
            return results
        except Exception:
            return []

    def collect_system_info(self):
        """采集基础系统信息"""
        self.profile.hostname = platform.node()
        self.profile.username = os.environ.get("USERNAME", "")
        self.profile.domain = os.environ.get("USERDOMAIN", "")
        self.profile.arch = platform.machine()
        self.profile.os_version = platform.platform()
        self.profile.os_build = platform.version()

        # 管理员权限检测
        try:
            self.profile.is_admin = bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            self.profile.is_admin = False

        # 内存信息
        try:
            import psutil
            mem = psutil.virtual_memory()
            self.profile.memory_mb = mem.total // (1024 * 1024)
        except ImportError:
            # fallback: ctypes
            try:
                class MEMORYSTATUSEX(ctypes.Structure):
                    _fields_ = [
                        ("dwLength", ctypes.c_ulong),
                        ("dwMemoryLoad", ctypes.c_ulong),
                        ("ullTotalPhys", ctypes.c_ulonglong),
                        ("ullAvailPhys", ctypes.c_ulonglong),
                        ("ullTotalPageFile", ctypes.c_ulonglong),
                        ("ullAvailPageFile", ctypes.c_ulonglong),
                        ("ullTotalVirtual", ctypes.c_ulonglong),
                        ("ullAvailVirtual", ctypes.c_ulonglong),
                        ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
                    ]
                mem = MEMORYSTATUSEX()
                mem.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
                ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(mem))
                self.profile.memory_mb = mem.ullTotalPhys // (1024 * 1024)
            except Exception:
                self.profile.memory_mb = 0

        # 磁盘信息
        try:
            import shutil
            _, _, free = shutil.disk_usage("C:\\")
            self.profile.disk_free_gb = round(free / (1024 ** 3), 1)
        except Exception:
            self.profile.disk_free_gb = 0.0

    def collect_processes(self):
        """采集进程列表并识别软件"""
        self.profile.running_processes = []
        self.profile.security_software = []
        self.profile.whitelist_software = []

        try:
            import subprocess
            output = subprocess.check_output(
                "tasklist /FO CSV /NH", shell=True, timeout=10
            ).decode("utf-8", errors="ignore")
            for line in output.strip().split("\n"):
                if line:
                    parts = line.split('","')
                    if parts:
                        proc_name = parts[0].strip('"').strip()
                        self.profile.running_processes.append(proc_name)
        except Exception:
            pass

        proc_lower = {p.lower() for p in self.profile.running_processes}

        # 识别安全软件
        for name, proc_list in SECURITY_SOFTWARE.items():
            for proc in proc_list:
                if proc.lower() in proc_lower:
                    self.profile.security_software.append(name)
                    break

        # 识别白名单软件
        for name, proc_list in WHITELIST_SOFTWARE.items():
            for proc in proc_list:
                if proc.lower() in proc_lower:
                    self.profile.whitelist_software.append(name)
                    break

    def run(self) -> EnvProfile:
        """运行完整环境感知"""
        self._init_wmi()
        self.collect_system_info()
        self.collect_processes()
        return self.profile


def perceive() -> EnvProfile:
    """便捷入口：运行环境感知"""
    ep = EnvPerception()
    return ep.run()


if __name__ == "__main__":
    import json
    profile = perceive()
    print(json.dumps(profile.to_dict(), indent=2, ensure_ascii=False))
