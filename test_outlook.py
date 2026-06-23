"""
SMTP/IMAP 邮件通道测试脚本
测试内容: SMTP 连接 → 发送加密邮件 → IMAP 接收解密 → 加解密往返
用法: python test_outlook.py <QQ邮箱> <授权码>
"""
import sys
import os
import time
import base64
import logging

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(name)s: %(message)s")
logger = logging.getLogger("test")


def test_smtp_connect(email: str, auth_code: str):
    """测试 1: SMTP 连接"""
    print("=" * 50)
    print("[1/4] SMTP 连接测试")
    print("=" * 50)

    from shared.crypto import CryptoEngine
    from agent.comm_channels import SMTPChannel

    crypto = CryptoEngine()
    ch = SMTPChannel(crypto, "test-smtp")
    ok = ch.connect(email=email, auth_code=auth_code)
    if ok:
        print(f"  [OK] SMTP 连接成功: {email}")
    else:
        print("  [FAIL] SMTP 连接失败 - 检查邮箱地址和授权码")
    ch.close()
    return ok


def test_smtp_send(email: str, auth_code: str):
    """测试 2: 发送加密邮件"""
    print("\n" + "=" * 50)
    print("[2/4] SMTP 发送测试")
    print("=" * 50)

    from shared.crypto import CryptoEngine
    from agent.comm_channels import SMTPChannel

    crypto = CryptoEngine()
    crypto.generate_rsa_keypair()
    # 自签公钥用于测试（实际场景由密钥交换获得）
    crypto.load_rsa_public(crypto.get_public_key_pem())
    ch = SMTPChannel(crypto, "test-send")
    if not ch.connect(email=email, auth_code=auth_code):
        print("  [FAIL] 连接失败")
        return False

    payload = b"SMTP channel test " + str(int(time.time())).encode()
    print(f"  [*] 发送: {payload.decode()}")
    ok = ch.send(payload, msg_type="result")
    if ok:
        print("  [OK] 邮件发送成功 (发给自己)")
    else:
        print("  [FAIL] 发送失败")
    ch.close()
    return ok


def test_imap_recv(email: str, auth_code: str):
    """测试 3: IMAP 接收"""
    print("\n" + "=" * 50)
    print("[3/4] IMAP 接收测试")
    print("=" * 50)

    from shared.crypto import CryptoEngine
    from agent.comm_channels import SMTPChannel

    crypto = CryptoEngine()
    crypto.generate_rsa_keypair()
    crypto.load_rsa_public(crypto.get_public_key_pem())
    ch = SMTPChannel(crypto, "test-recv")
    if not ch.connect(email=email, auth_code=auth_code):
        print("  [FAIL] 连接失败")
        return False

    print("  [*] 扫描未读邮件...")
    data = ch.recv()
    if data:
        print(f"  [OK] 收到: {data}")
        ch.close()
        return True
    else:
        print("  [INFO] 无可解析的未读邮件 (先运行发送测试)")
        ch.close()
        return True


def test_encrypt_roundtrip():
    """测试 4: 加解密往返 (不依赖网络)"""
    print("\n" + "=" * 50)
    print("[4/4] 加解密往返")
    print("=" * 50)

    from shared.crypto import CryptoEngine, serialize_message, deserialize_message

    # 模拟双方交换公钥
    sender = CryptoEngine()
    sender.generate_rsa_keypair()
    receiver = CryptoEngine()
    receiver.generate_rsa_keypair()
    sender.load_rsa_public(receiver.get_public_key_pem())
    receiver.load_rsa_public(sender.get_public_key_pem())

    original = b"roundtrip-" + os.urandom(32)
    print(f"  [*] 原始: {len(original)} bytes")

    msg = sender.encrypt_message(original, "test", channel="smtp", msg_type="exfil")
    wire = serialize_message(msg)
    encoded = base64.b64encode(wire).decode("ascii")
    print(f"  [OK] 加密+base64: {len(encoded)} chars")

    wire_back = base64.b64decode(encoded)
    msg_back, _ = deserialize_message(wire_back)
    result = receiver.decrypt_message(msg_back)

    if result == original:
        print("  [OK] 数据一致!")
        return True
    print("  [FAIL] 数据不一致")
    return False


def main():
    if len(sys.argv) >= 3:
        email, auth_code = sys.argv[1], sys.argv[2]
    else:
        print("用法: python test_outlook.py <QQ邮箱> <授权码>")
        print("示例: python test_outlook.py 123456@qq.com abcdefghijklmn")
        print()
        print("获取授权码: QQ邮箱 → 设置 → 账户 → POP3/IMAP → 生成授权码")
        return

    results = {}
    results["SMTP 连接"] = test_smtp_connect(email, auth_code)
    results["加解密往返"] = test_encrypt_roundtrip()

    if results["SMTP 连接"]:
        results["邮件发送"] = test_smtp_send(email, auth_code)
        if results["邮件发送"]:
            print("\n  等待 5 秒...")
            time.sleep(5)
        results["邮件接收"] = test_imap_recv(email, auth_code)

    print("\n" + "=" * 50)
    print("结果汇总")
    print("=" * 50)
    for name, ok in results.items():
        print(f"  [{'PASS' if ok else 'FAIL'}] {name}")
    passed = sum(v for v in results.values())
    print(f"\n  {passed}/{len(results)} 通过")


if __name__ == "__main__":
    main()
