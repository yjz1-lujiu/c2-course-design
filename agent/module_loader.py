"""
sys.meta_path 动态模块加载器
从加密文件/内网共享目录加载插件，内存中解密执行，无磁盘落地
"""
import sys
import os
import marshal
import types
import importlib.abc
import importlib.machinery


class EncryptedModuleFinder(importlib.abc.MetaPathFinder):
    """
    自定义 MetaPathFinder
    拦截特定前缀的 import 请求，从加密插件目录加载
    """

    def __init__(self, plugin_dir: str, crypto_engine, prefix: str = "plugins."):
        """
        Args:
            plugin_dir: 加密插件文件所在目录
            crypto_engine: 共享加密引擎实例 (用于 AES 解密)
            prefix: 模块名前缀，匹配的 import 才会被拦截
        """
        self.plugin_dir = plugin_dir
        self.crypto_engine = crypto_engine
        self.prefix = prefix
        self._manifest = {}

    def load_manifest(self):
        """加载加密的插件清单"""
        manifest_path = os.path.join(self.plugin_dir, "__enc_manifest__")
        if not os.path.exists(manifest_path):
            self._manifest = {}
            return

        with open(manifest_path, "rb") as f:
            data = f.read()

        # 清单格式: [4字节nonce长度][nonce][4字节tag长度][tag][密文]
        import struct
        offset = 0
        nonce_len = struct.unpack_from("!I", data, offset)[0]
        offset += 4
        nonce = data[offset:offset + nonce_len]
        offset += nonce_len
        tag_len = struct.unpack_from("!I", data, offset)[0]
        offset += 4
        tag = data[offset:offset + tag_len]
        offset += tag_len
        ciphertext = data[offset:]

        plaintext = self.crypto_engine.aes_decrypt(nonce, ciphertext, tag)
        import json
        self._manifest = json.loads(plaintext.decode("utf-8"))

    def find_module(self, fullname, path=None):
        """匹配特定前缀的模块名"""
        if fullname.startswith(self.prefix):
            # 从 manifest 中查找
            module_key = fullname[len(self.prefix):]
            if module_key in self._manifest or self._try_find_file(module_key):
                return self
        return None

    def _try_find_file(self, module_key: str) -> bool:
        """检查加密插件文件是否存在"""
        enc_path = os.path.join(self.plugin_dir, f"{module_key}.enc")
        return os.path.exists(enc_path)

    def load_module(self, fullname):
        """加载并解密模块"""
        if fullname in sys.modules:
            return sys.modules[fullname]

        module_key = fullname[len(self.prefix):]
        enc_path = os.path.join(self.plugin_dir, f"{module_key}.enc")

        if not os.path.exists(enc_path):
            raise ImportError(f"Encrypted plugin not found: {module_key}")

        # 读取加密文件
        with open(enc_path, "rb") as f:
            data = f.read()

        # 解析: [4字节nonce长度][nonce][4字节tag长度][tag][密文]
        import struct
        offset = 0
        nonce_len = struct.unpack_from("!I", data, offset)[0]
        offset += 4
        nonce = data[offset:offset + nonce_len]
        offset += nonce_len
        tag_len = struct.unpack_from("!I", data, offset)[0]
        offset += 4
        tag = data[offset:offset + tag_len]
        offset += tag_len
        ciphertext = data[offset:]

        # AES 解密得到 .pyc 字节码
        pyc_data = self.crypto_engine.aes_decrypt(nonce, ciphertext, tag)

        # 跳过 .pyc 头部 (magic number + timestamp + size = 16 bytes in Python 3.8+)
        pyc_header_len = 16
        code = marshal.loads(pyc_data[pyc_header_len:])

        # 创建模块对象
        module = types.ModuleType(fullname)
        module.__file__ = f"<encrypted:{module_key}>"
        module.__loader__ = self
        module.__name__ = fullname
        module.__package__ = self.prefix.rstrip(".")

        # 将模块注册到 sys.modules 再执行（允许模块自引用）
        sys.modules[fullname] = module

        try:
            exec(code, module.__dict__)
        except Exception:
            del sys.modules[fullname]
            raise

        return module


class EncryptedModuleLoader(importlib.abc.Loader):
    """备用 Loader（目前 find_module/load_module 一体化，此类为扩展预留）"""
    pass


def install_finder(plugin_dir: str, crypto_engine, prefix: str = "plugins.") -> EncryptedModuleFinder:
    """
    安装加密模块 finder 到 sys.meta_path

    用法:
        finder = install_finder("./plugins", crypto_engine)
        finder.load_manifest()

        # 之后就可以 import plugins.xxx
        from plugins import recon  # 自动解密加载
    """
    finder = EncryptedModuleFinder(plugin_dir, crypto_engine, prefix)
    # 移除旧的同类型 finder（避免重复注册）
    sys.meta_path = [f for f in sys.meta_path
                     if not isinstance(f, EncryptedModuleFinder)]
    sys.meta_path.insert(0, finder)
    return finder


def uninstall_finder():
    """移除加密模块 finder"""
    sys.meta_path = [f for f in sys.meta_path
                     if not isinstance(f, EncryptedModuleFinder)]


def load_plugin_direct(plugin_dir: str, module_name: str, crypto_engine) -> types.ModuleType:
    """
    直接从加密文件加载单个插件模块（不通过 import 机制）

    Args:
        plugin_dir: 插件目录
        module_name: 模块名（不含 .enc 后缀）
        crypto_engine: 加密引擎

    Returns:
        加载后的模块对象
    """
    enc_path = os.path.join(plugin_dir, f"{module_name}.enc")
    if not os.path.exists(enc_path):
        raise FileNotFoundError(f"Plugin not found: {enc_path}")

    with open(enc_path, "rb") as f:
        data = f.read()

    import struct
    offset = 0
    nonce_len = struct.unpack_from("!I", data, offset)[0]
    offset += 4
    nonce = data[offset:offset + nonce_len]
    offset += nonce_len
    tag_len = struct.unpack_from("!I", data, offset)[0]
    offset += 4
    tag = data[offset:offset + tag_len]
    offset += tag_len
    ciphertext = data[offset:]

    pyc_data = crypto_engine.aes_decrypt(nonce, ciphertext, tag)
    code = marshal.loads(pyc_data[16:])

    module = types.ModuleType(module_name)
    module.__file__ = f"<encrypted:{module_name}>"
    exec(code, module.__dict__)
    return module
