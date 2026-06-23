"""
示例插件：内网横向探测模块 (拓展功能)
自动发现同网段存活主机
"""
import socket
import subprocess
import struct
import threading
import time
import json


def get_local_ip() -> str:
    """获取本机 IP"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def ping_host(ip: str, timeout: float = 1.0) -> bool:
    """使用系统 ping 检测主机存活"""
    try:
        result = subprocess.run(
            ["ping", "-n", "1", "-w", str(int(timeout * 1000)), ip],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=timeout + 1,
            creationflags=0x08000000
        )
        return result.returncode == 0
    except Exception:
        return False


def tcp_probe(ip: str, port: int = 445, timeout: float = 1.0) -> bool:
    """TCP 端口探测（SMB 445 端口通常在 Windows 主机上开放）"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(timeout)
        result = s.connect_ex((ip, port))
        s.close()
        return result == 0
    except Exception:
        return False


def resolve_hostname(ip: str) -> str:
    """反向 DNS 解析"""
    try:
        return socket.gethostbyaddr(ip)[0]
    except Exception:
        return ""


def scan_subnet(subnet_prefix: str = None, start: int = 1, end: int = 254,
                max_threads: int = 50) -> list:
    """
    扫描子网内存活主机

    Args:
        subnet_prefix: 网段前缀 (如 "192.168.1"), 自动检测
        start: 起始主机号
        end: 结束主机号
        max_threads: 最大并发线程数

    Returns:
        存活主机列表
    """
    if not subnet_prefix:
        local_ip = get_local_ip()
        subnet_prefix = ".".join(local_ip.split(".")[:3])

    alive_hosts = []
    lock = threading.Lock()
    sem = threading.Semaphore(max_threads)

    def scan_one(host_id: int):
        ip = f"{subnet_prefix}.{host_id}"
        sem.acquire()
        try:
            # 先尝试 TCP 445 (更快)
            if tcp_probe(ip, 445, timeout=0.5):
                hostname = resolve_hostname(ip)
                with lock:
                    alive_hosts.append({
                        "ip": ip,
                        "hostname": hostname,
                        "method": "tcp:445",
                    })
            elif ping_host(ip, timeout=0.5):
                hostname = resolve_hostname(ip)
                with lock:
                    alive_hosts.append({
                        "ip": ip,
                        "hostname": hostname,
                        "method": "icmp",
                    })
        finally:
            sem.release()

    threads = []
    for host_id in range(start, end + 1):
        t = threading.Thread(target=scan_one, args=(host_id,))
        t.start()
        threads.append(t)

    for t in threads:
        t.join(timeout=5)

    return sorted(alive_hosts, key=lambda x: tuple(int(p) for p in x["ip"].split(".")))


def get_arp_table() -> list:
    """从 ARP 表提取已知主机"""
    try:
        output = subprocess.check_output(
            "arp -a", shell=True, timeout=5,
            creationflags=0x08000000
        ).decode("utf-8", errors="ignore")

        entries = []
        for line in output.split("\n"):
            line = line.strip()
            parts = line.split()
            if len(parts) >= 3 and "." in parts[0]:
                entries.append({
                    "ip": parts[0],
                    "mac": parts[1],
                    "type": parts[2] if len(parts) > 2 else "unknown",
                })
        return entries
    except Exception:
        return []


def run(subnet: str = None) -> dict:
    """插件入口"""
    result = {
        "local_ip": get_local_ip(),
        "arp_table": get_arp_table(),
        "alive_hosts": [],
    }

    # 扫描子网
    alive = scan_subnet(subnet)
    result["alive_hosts"] = alive
    result["total_alive"] = len(alive)

    return result


if __name__ == "__main__":
    print(json.dumps(run(), indent=2, ensure_ascii=False))
