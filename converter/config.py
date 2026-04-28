"""配置管理模块 — 负责应用配置的加载、保存和默认值"""

import os
import sys
import json
import base64

from .generator import SCRIPT_DIR
from .logging_config import get_logger

logger = get_logger('config')


def _dpapi_encrypt(data_bytes):
    """使用 Windows DPAPI 加密数据"""
    if sys.platform != 'win32':
        return base64.b64encode(data_bytes).decode('ascii')
    try:
        import ctypes
        import ctypes.wintypes

        class DATA_BLOB(ctypes.Structure):
            _fields_ = [('cbData', ctypes.wintypes.DWORD),
                         ('pbData', ctypes.POINTER(ctypes.c_char))]

        blob_in = DATA_BLOB(len(data_bytes), ctypes.create_string_buffer(data_bytes, len(data_bytes)))
        blob_out = DATA_BLOB()
        if ctypes.windll.crypt32.CryptProtectData(
            ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out),
        ):
            enc = ctypes.string_at(blob_out.pbData, blob_out.cbData)
            ctypes.windll.kernel32.LocalFree(blob_out.pbData)
            return base64.b64encode(enc).decode('ascii')
    except Exception:
        logger.debug('DPAPI 加密失败，回退到 base64', exc_info=True)
    return base64.b64encode(data_bytes).decode('ascii')


def _dpapi_decrypt(encoded_str):
    """使用 Windows DPAPI 解密数据"""
    raw = base64.b64decode(encoded_str)
    if sys.platform != 'win32':
        return raw.decode('utf-8')
    try:
        import ctypes
        import ctypes.wintypes

        class DATA_BLOB(ctypes.Structure):
            _fields_ = [('cbData', ctypes.wintypes.DWORD),
                         ('pbData', ctypes.POINTER(ctypes.c_char))]

        blob_in = DATA_BLOB(len(raw), ctypes.create_string_buffer(raw, len(raw)))
        blob_out = DATA_BLOB()
        if ctypes.windll.crypt32.CryptUnprotectData(
            ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out),
        ):
            dec = ctypes.string_at(blob_out.pbData, blob_out.cbData)
            ctypes.windll.kernel32.LocalFree(blob_out.pbData)
            return dec.decode('utf-8')
    except Exception:
        logger.debug('DPAPI 解密失败，尝试 base64 回退', exc_info=True)
    return raw.decode('utf-8')


def encrypt_api_key(key):
    """加密 API Key"""
    if not key:
        return ''
    return 'ENC:' + _dpapi_encrypt(key.encode('utf-8'))


def decrypt_api_key(stored):
    """解密 API Key，兼容明文存储的旧格式"""
    if not stored:
        return ''
    if stored.startswith('ENC:'):
        try:
            return _dpapi_decrypt(stored[4:])
        except Exception:
            logger.warning('API Key 解密失败', exc_info=True)
            return ''
    return stored

if getattr(sys, 'frozen', False):
    CONFIG_PATH = os.path.join(os.path.dirname(sys.executable), 'config.json')
else:
    CONFIG_PATH = os.path.join(SCRIPT_DIR, 'config.json')

APP_VERSION = '3.0.0'

AI_MODELS = [
    'gpt-5.5',
]

DEFAULT_CONFIG = {
    'api_key': '',
    'default_template': '默认',
    'output_dir': '',
    'batch_output_dir': '',
    'last_save_dir': '',
    'ai_model': 'gpt-5.5',
    'theme': 'light',
    'recent_files': [],
    'window_geometry': '',
    'export_history': [],
    'doc_features': {
        'toc_enabled': False,
        'header_enabled': False,
        'header_text': '',
        'footer_enabled': True,
        'page_number': True,
        'logo_path': '',
        'watermark_enabled': False,
        'watermark_text': '',
    },
    'format_settings': {
        'preset': '默认',
        'custom_overrides': {},
        'docx_template': '',
    },
}

THEMES = {
    'light': {
        'bg': '#ffffff',
        'fg': '#000000',
        'insert_bg': '#000000',
        'select_bg': '#b3d9ff',
        'preview_bg': '#fafafa',
        'code_bg': '#f0f0f0',
        'heading_fg': '#1a56db',
        'link_fg': '#0563C1',
        'bold_fg': '#000000',
        'task_fg': '#228B22',
    },
    'dark': {
        'bg': '#1e1e1e',
        'fg': '#d4d4d4',
        'insert_bg': '#d4d4d4',
        'select_bg': '#264f78',
        'preview_bg': '#252526',
        'code_bg': '#2d2d2d',
        'heading_fg': '#569cd6',
        'link_fg': '#4fc1ff',
        'bold_fg': '#dcdcaa',
        'task_fg': '#6a9955',
    },
}

MAX_RECENT = 10
MAX_HISTORY = 50
DRAFT_FILENAME = '.md_tong_draft.md'

MARKDOWN_CHEATSHEET = """
# Markdown 语法速查

## 标题
# 一级标题
## 二级标题
### 三级标题

## 文本格式
**加粗文本**
*斜体文本*
~~删除线~~
`行内代码`

## 列表
- 无序列表项
1. 有序列表项
- [ ] 未完成任务
- [x] 已完成任务

## 链接与图片
[链接文字](URL)
![图片描述](图片路径)

## 表格
| 表头1 | 表头2 |
|-------|-------|
| 单元格 | 单元格 |

## 代码块
```python
print("Hello")
```

## 引用
> 引用文本

## 分隔线
---
""".strip()


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    if os.path.isfile(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                saved = json.load(f)
            for k, v in saved.items():
                if k == 'doc_features' and isinstance(v, dict):
                    cfg['doc_features'] = {**DEFAULT_CONFIG['doc_features'], **v}
                elif k == 'format_settings' and isinstance(v, dict):
                    cfg['format_settings'] = {**DEFAULT_CONFIG['format_settings'], **v}
                else:
                    cfg[k] = v
        except Exception:
            logger.warning('配置文件加载失败，使用默认配置', exc_info=True)
    cfg['api_key'] = decrypt_api_key(cfg.get('api_key', ''))
    return cfg


def save_config(cfg):
    data = dict(cfg)
    data['api_key'] = encrypt_api_key(data.get('api_key', ''))
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
