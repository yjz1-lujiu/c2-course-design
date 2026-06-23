"""
插件打包加密工具
将 .py 文件编译为 .pyc → AES 加密 → 存入 plugins/ 目录
生成加密清单 manifest
"""
import os
import sys
import json
import struct
import marshal
import compileall
import tempfile
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from shared.crypto import CryptoEngine


def compile_to_pyc(py_path: str) -> bytes:
    """将 .py 文件编译为 .pyc 字节码"""
    import py_compile
    with tempfile.NamedTemporaryFile(suffix=".pyc", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        py_compile.compile(py_path, cfile=tmp_path, doraise=True)
        with open(tmp_path, "rb") as f:
            return f.read()
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass


def encrypt_plugin(pyc_data: bytes, crypto: CryptoEngine) -> bytes:
    """
    加密 .pyc 数据
    格式: [4字节nonce长度][nonce][4字节tag长度][tag][AES-GCM密文]
    """
    nonce, ciphertext, tag = crypto.aes_encrypt(pyc_data)

    parts = []
    parts.append(struct.pack("!I", len(nonce)))
    parts.append(nonce)
    parts.append(struct.pack("!I", len(tag)))
    parts.append(tag)
    parts.append(ciphertext)

    return b"".join(parts)


def pack_single(py_path: str, output_dir: str, crypto: CryptoEngine) -> str:
    """打包单个插件文件"""
    py_name = os.path.splitext(os.path.basename(py_path))[0]

    # 编译
    print(f"  Compiling: {py_path}")
    pyc_data = compile_to_pyc(py_path)

    # 加密
    print(f"  Encrypting: {py_name}")
    enc_data = encrypt_plugin(pyc_data, crypto)

    # 写入 .enc 文件
    enc_path = os.path.join(output_dir, f"{py_name}.enc")
    with open(enc_path, "wb") as f:
        f.write(enc_data)

    print(f"  Output: {enc_path} ({len(enc_data)} bytes)")
    return py_name


def build_manifest(plugin_dir: str, plugin_names: list, crypto: CryptoEngine):
    """生成加密的插件清单"""
    manifest = {
        "plugins": {name: {"version": "1.0"} for name in plugin_names},
        "created": __import__("time").time(),
    }

    manifest_json = json.dumps(manifest, ensure_ascii=False).encode("utf-8")
    nonce, ciphertext, tag = crypto.aes_encrypt(manifest_json)

    parts = []
    parts.append(struct.pack("!I", len(nonce)))
    parts.append(nonce)
    parts.append(struct.pack("!I", len(tag)))
    parts.append(tag)
    parts.append(ciphertext)

    manifest_path = os.path.join(plugin_dir, "__enc_manifest__")
    with open(manifest_path, "wb") as f:
        f.write(b"".join(parts))

    print(f"  Manifest: {manifest_path}")


def main():
    parser = argparse.ArgumentParser(description="Plugin packer - compile and encrypt Python plugins")
    parser.add_argument("input", nargs="+", help="Plugin .py files or directory containing .py files")
    parser.add_argument("-o", "--output", default="agent/plugins", help="Output directory")
    parser.add_argument("-k", "--key", help="AES key (hex, 32 bytes). Auto-generated if omitted.")
    parser.add_argument("--key-file", help="File to save the generated AES key")
    args = parser.parse_args()

    # 初始化加密引擎
    crypto = CryptoEngine()
    if args.key:
        key = bytes.fromhex(args.key)
        crypto.session_key = key
        print(f"[*] Using provided key: {args.key[:16]}...")
    else:
        crypto.generate_session_key()
        key_hex = crypto.session_key.hex()
        print(f"[*] Generated key: {key_hex[:16]}...")
        if args.key_file:
            with open(args.key_file, "w") as f:
                f.write(key_hex)
            print(f"[*] Key saved to: {args.key_file}")

    # 确保输出目录存在
    os.makedirs(args.output, exist_ok=True)

    # 收集所有 .py 文件
    py_files = []
    for item in args.input:
        if os.path.isdir(item):
            for fname in os.listdir(item):
                if fname.endswith(".py") and not fname.startswith("__"):
                    py_files.append(os.path.join(item, fname))
        elif os.path.isfile(item) and item.endswith(".py"):
            py_files.append(item)
        else:
            print(f"[!] Skipping: {item}")

    if not py_files:
        print("[!] No .py files found.")
        return

    # 打包每个插件
    print(f"\n[*] Packing {len(py_files)} plugin(s) to {args.output}/")
    plugin_names = []
    for py_path in py_files:
        name = pack_single(py_path, args.output, crypto)
        plugin_names.append(name)

    # 生成清单
    print(f"\n[*] Building manifest...")
    build_manifest(args.output, plugin_names, crypto)

    print(f"\n[+] Done! {len(plugin_names)} plugin(s) packed.")
    print(f"    Key (hex): {crypto.session_key.hex()}")


if __name__ == "__main__":
    main()
