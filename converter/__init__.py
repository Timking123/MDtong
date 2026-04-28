"""converter 包 — 会议纪要转换工具核心模块"""

import os
import sys

from .logging_config import setup_logging, get_logger
from .config import load_config, save_config
from .parser import parse_markdown, parse_inline_bold, parse_inline, parse_front_matter
from .generator import convert, auto_filename, SCRIPT_DIR
from .gui import run_gui
from .templates import list_templates, load_template_content
from .reverse import docx_to_markdown
from .ai_polish import polish
from .pdf_export import export_pdf
from .html_export import export_html

setup_logging()


def get_resource_dir():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


__all__ = [
    'parse_markdown',
    'parse_front_matter',
    'parse_inline_bold',
    'parse_inline',
    'convert',
    'auto_filename',
    'run_gui',
    'list_templates',
    'load_template_content',
    'docx_to_markdown',
    'polish',
    'export_pdf',
    'export_html',
    'load_config',
    'save_config',
    'get_resource_dir',
]
