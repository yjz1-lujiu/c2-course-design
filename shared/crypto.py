import os
import json
import time
import struct
from Crypto.Cipher import AES, PKCS1_OAEP
from Crypto.PublicKey import RSA
from Crypto.Random import get_random_bytes


class CryptoEngine:
    """AES-256-GCM + RSA-2048 混合加密引擎"""

    def wipe(self):
        """
        显式清零内存中的密钥材料
        将所有密钥引用置为 None，使 Python GC 尽快回收
        对 session_key 使用 mutable bytearray 覆写原始字节
        """
        # 覆写 session_key 内存（bytearray 是 mutable 的，覆写后 GC 回收）
        if self.session_key is not None:
            try:
                # 创建同长度全零 bytearray 覆写
                key_len = len(self.session_key)
                zero_buf = bytearray(key_len)
                self.session_key = bytes(zero_buf)
            except Exception:
                pass
            self.session_key = None

        # RSA 密钥对象内部包含大整数，置为 None 让 GC 回收
        if self.rsa_key is not None:
            try:
                self.rsa_key._key = None
            except Exception:
                pass
            self.rsa_key = None

        if self.peer_public_key is not None:
            try:
                self.peer_public_key._key = None
            except Exception:
                pass
            self.peer_public_key = None

        self._seq = 0

    def __del__(self):
        """析构时自动清零密钥材料"""
        self.wipe()

    def __init__(self):
        self.rsa_key = None
        self.peer_public_key = None
        self.session_key = None
        self._seq = 0

    # ── RSA 密钥管理 ──────────────────────────────────────

    def generate_rsa_keypair(self, bits=2048):
        """生成 RSA 密钥对"""
        self.rsa_key = RSA.generate(bits)
        return self.rsa_key

    def load_rsa_private(self, key_data: bytes):
        """从 PEM/DER 加载私钥"""
        self.rsa_key = RSA.import_key(key_data)

    def load_rsa_public(self, key_data: bytes):
        """加载对方公钥"""
        self.peer_public_key = RSA.import_key(key_data)

    def get_public_key_pem(self) -> bytes:
        """导出本方公钥 PEM"""
        if self.rsa_key is None:
            raise ValueError("RSA key not generated")
        return self.rsa_key.publickey().export_key()

    def get_private_key_pem(self) -> bytes:
        """导出本方私钥 PEM"""
        if self.rsa_key is None:
            raise ValueError("RSA key not generated")
        return self.rsa_key.export_key()

    # ── 会话密钥 ──────────────────────────────────────────

    def generate_session_key(self) -> bytes:
        """生成随机会话密钥 (AES-256)"""
        self.session_key = get_random_bytes(32)
        self._seq = 0
        return self.session_key

    def encrypt_session_key(self, session_key: bytes = None) -> bytes:
        """用对方 RSA 公钥加密会话密钥"""
        key = session_key or self.session_key
        if self.peer_public_key is None:
            raise ValueError("Peer public key not loaded")
        cipher = PKCS1_OAEP.new(self.peer_public_key)
        return cipher.encrypt(key)

    def decrypt_session_key(self, encrypted_key: bytes) -> bytes:
        """用本方 RSA 私钥解密会话密钥"""
        if self.rsa_key is None:
            raise ValueError("RSA private key not loaded")
        cipher = PKCS1_OAEP.new(self.rsa_key)
        self.session_key = cipher.decrypt(encrypted_key)
        self._seq = 0
        return self.session_key

    # ── AES-256-GCM 加解密 ────────────────────────────────

    def aes_encrypt(self, plaintext: bytes, key: bytes = None) -> tuple:
        """
        AES-256-GCM 加密
        返回: (nonce, ciphertext, tag)
        """
        k = key or self.session_key
        if k is None:
            raise ValueError("Session key not set")
        cipher = AES.new(k, AES.MODE_GCM)
        ciphertext, tag = cipher.encrypt_and_digest(plaintext)
        return cipher.nonce, ciphertext, tag

    def aes_decrypt(self, nonce: bytes, ciphertext: bytes, tag: bytes,
                    key: bytes = None) -> bytes:
        """AES-256-GCM 解密"""
        k = key or self.session_key
        if k is None:
            raise ValueError("Session key not set")
        cipher = AES.new(k, AES.MODE_GCM, nonce=nonce)
        return cipher.decrypt_and_verify(ciphertext, tag)

    # ── 完整消息加密/解密 ──────────────────────────────────

    def encrypt_message(self, plaintext: bytes, session_id: str,
                        channel: str = "tcp", msg_type: str = "cmd") -> dict:
        """
        加密完整消息，返回可序列化的字典结构

        消息结构:
        {
            "header": {session_id, seq, timestamp, channel, type},
            "encrypted_key": bytes,   # RSA 加密的会话密钥
            "nonce": bytes,           # GCM nonce
            "payload": bytes,         # AES-GCM 加密的业务数据
            "tag": bytes              # GCM 认证标签
        }
        """
        if self.session_key is None:
            self.generate_session_key()

        nonce, ciphertext, tag = self.aes_encrypt(plaintext)
        enc_key = self.encrypt_session_key()

        self._seq += 1
        header = {
            "session_id": session_id,
            "seq": self._seq,
            "timestamp": time.time(),
            "channel": channel,
            "type": msg_type,
        }

        return {
            "header": header,
            "encrypted_key": enc_key,
            "nonce": nonce,
            "payload": ciphertext,
            "tag": tag,
        }

    def decrypt_message(self, message: dict, max_age: float = 300.0) -> bytes:
        """
        解密消息，同时验证时间戳防重放

        max_age: 消息最大有效时间（秒），默认5分钟
        """
        header = message["header"]
        age = abs(time.time() - header["timestamp"])
        if age > max_age:
            raise ValueError(f"Message expired (age={age:.1f}s, max={max_age}s)")

        # 解密会话密钥
        self.decrypt_session_key(message["encrypted_key"])

        # 解密业务数据
        plaintext = self.aes_decrypt(
            message["nonce"],
            message["payload"],
            message["tag"],
        )
        return plaintext


def serialize_message(msg: dict) -> bytes:
    """将消息字典序列化为字节流（网络传输用）"""
    import base64
    wire = {
        "header": msg["header"],
        "encrypted_key": base64.b64encode(msg["encrypted_key"]).decode(),
        "nonce": base64.b64encode(msg["nonce"]).decode(),
        "payload": base64.b64encode(msg["payload"]).decode(),
        "tag": base64.b64encode(msg["tag"]).decode(),
    }
    data = json.dumps(wire, ensure_ascii=False).encode("utf-8")
    # 4字节长度前缀 + JSON 数据
    return struct.pack("!I", len(data)) + data


def deserialize_message(raw: bytes, offset: int = 0) -> tuple:
    """
    从字节流反序列化消息
    返回: (message_dict, consumed_bytes)
    """
    import base64
    if len(raw) - offset < 4:
        raise ValueError("Not enough data for length prefix")
    length = struct.unpack("!I", raw[offset:offset + 4])[0]
    data = raw[offset + 4:offset + 4 + length]
    if len(data) < length:
        raise ValueError("Incomplete message data")

    wire = json.loads(data.decode("utf-8"))
    msg = {
        "header": wire["header"],
        "encrypted_key": base64.b64decode(wire["encrypted_key"]),
        "nonce": base64.b64decode(wire["nonce"]),
        "payload": base64.b64decode(wire["payload"]),
        "tag": base64.b64decode(wire["tag"]),
    }
    return msg, 4 + length
