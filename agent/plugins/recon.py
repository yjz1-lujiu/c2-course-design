"""
示例插件：网络侦察模块
通过 Agent 动态加载执行内网信息收集
"""
import os
import re
import socket
import subprocess
import json


def _run_cmd(cmd: str) -> str:
    try:
        raw = subprocess.check_output(
            cmd, shell=True, timeout=10,
            creationflags=0x08000000
        )
        # 中文 Windows 命令输出是 GBK，优先用 GBK 解码
        for enc in ("gbk", "utf-8", "gb2312"):
            try:
                return raw.decode(enc)
            except (UnicodeDecodeError, LookupError):
                continue
        return raw.decode("utf-8", errors="replace")
    except Exception:
        return ""


def _parse_interfaces(raw: str) -> list:
    """从 ipconfig /all 解析网卡信息（兼容中英文 Windows）"""
    adapters = []
    current = {}
    for line in raw.splitlines():
        line = line.strip()
        if not line:
            continue
        if not line.startswith(" ") and ":" not in line and line:
            if current:
                adapters.append(current)
            current = {"name": line.rstrip(".")}
        # IPv4 地址 (英文: IPv4 Address / 中文: IPv4 地址)
        elif ("IPv4" in line or "IP Address" in line or "IPv4 地址" in line) and ":" in line:
            m = re.search(r"(\d+\.\d+\.\d+\.\d+)", line)
            if m:
                current["ipv4"] = m.group(1)
        # 子网掩码
        elif ("Subnet Mask" in line or "子网掩码" in line) and ":" in line:
            m = re.search(r"(\d+\.\d+\.\d+\.\d+)", line)
            if m:
                current["mask"] = m.group(1)
        # 默认网关
        elif ("Default Gateway" in line or "默认网关" in line) and ":" in line:
            m = re.search(r"(\d+\.\d+\.\d+\.\d+)", line)
            if m:
                current["gateway"] = m.group(1)
        # MAC 地址
        elif ("Physical Address" in line or "物理地址" in line) and ":" in line:
            m = re.search(r"([0-9A-Fa-f]{2}(:[0-9A-Fa-f]{2}){5})", line)
            if m:
                current["mac"] = m.group(1)
        # DNS 服务器
        elif ("DNS Servers" in line or "DNS 服务器" in line) and ":" in line:
            m = re.search(r"(\d+\.\d+\.\d+\.\d+)", line)
            if m:
                current.setdefault("dns", []).append(m.group(1))
    if current:
        adapters.append(current)
    return adapters


def _parse_arp(raw: str) -> list:
    """解析 ARP 表"""
    entries = []
    for line in raw.splitlines():
        m = re.match(r"\s*(\d+\.\d+\.\d+\.\d+)\s+([0-9a-fA-F-]{17}|incomplete)\s+(\S+)", line)
        if m:
            entries.append({"ip": m.group(1), "mac": m.group(2), "type": m.group(3)})
    return entries


def _parse_netstat(raw: str) -> list:
    """解析 netstat，仅保留 LISTENING 和 ESTABLISHED"""
    conns = []
    for line in raw.splitlines():
        parts = line.split()
        if len(parts) >= 4 and parts[0] in ("TCP", "UDP"):
            state = parts[3] if parts[0] == "TCP" else ""
            if state in ("LISTENING", "ESTABLISHED", ""):
                conns.append({
                    "proto": parts[0],
                    "local": parts[1],
                    "remote": parts[2] if len(parts) > 2 else "",
                    "state": state,
                    "pid": parts[-1] if parts[-1].isdigit() else "",
                })
    return conns


def _parse_dns_cache(raw: str) -> list:
    """解析 DNS 缓存记录（兼容中英文 Windows）"""
    records = []
    current = {}
    for line in raw.splitlines():
        line = line.strip()
        # 兼容中文: 记录名 / 英文: Record Name
        if line.startswith("Record Name") or line.startswith("记录名"):
            if current:
                records.append(current)
            current = {"name": line.split(":", 1)[-1].strip()}
        elif (line.startswith("Record Type") or line.startswith("记录类型")) and current:
            current["type"] = line.split(":", 1)[-1].strip()
        elif (line.startswith("Record Data") or line.startswith("记录数据")) and current:
            current["data"] = line.split(":", 1)[-1].strip()
    if current:
        records.append(current)
    return records[:50]


def run(target_subnet: str = None) -> dict:
    """
    执行内网侦察，返回结构化结果
    """
    result = {
        "hostname": socket.gethostname(),
        "interfaces": _parse_interfaces(_run_cmd("ipconfig /all")),
        "arp_table": _parse_arp(_run_cmd("arp -a")),
        "dns_cache": _parse_dns_cache(_run_cmd("ipconfig /displaydns")),
        "connections": _parse_netstat(_run_cmd("netstat -ano")),
    }

    # 统计摘要
    result["summary"] = {
        "total_interfaces": len(result["interfaces"]),
        "active_interfaces": [a["name"] for a in result["interfaces"] if a.get("ipv4")],
        "arp_hosts": len(result["arp_table"]),
        "dns_records": len(result["dns_cache"]),
        "active_connections": len([c for c in result["connections"] if c["state"] == "ESTABLISHED"]),
        "listening_ports": len([c for c in result["connections"] if c["state"] == "LISTENING"]),
    }

    return result


if __name__ == "__main__":
    print(json.dumps(run(), indent=2, ensure_ascii=False))
