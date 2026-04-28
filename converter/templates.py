"""模板管理 — 扫描和加载 Markdown 内容模板"""

import os
import sys

if getattr(sys, 'frozen', False):
    TEMPLATES_DIR = os.path.join(sys._MEIPASS, 'templates')
else:
    TEMPLATES_DIR = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        'templates',
    )


def list_templates():
    """返回 templates/ 目录下所有 .md 文件名（不含扩展名）"""
    if not os.path.isdir(TEMPLATES_DIR):
        return []
    return [
        os.path.splitext(f)[0]
        for f in sorted(os.listdir(TEMPLATES_DIR))
        if f.lower().endswith('.md')
    ]


def load_template_content(name):
    """读取模板 .md 文件，返回其文本内容"""
    path = os.path.join(TEMPLATES_DIR, f'{name}.md')
    if not os.path.isfile(path):
        return None
    for enc in ('utf-8-sig', 'utf-8', 'gbk'):
        try:
            with open(path, 'r', encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    return None
