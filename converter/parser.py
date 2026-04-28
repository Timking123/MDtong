"""Markdown 解析器 — 将文本拆分为结构化 block 列表"""

import re
import yaml
from datetime import datetime


# ── 正则常量 ──

RE_EMPTY = re.compile(r'^\s*$')
RE_HR = re.compile(r'^\s*([-*_]\s*){3,}$')
RE_TABLE_SEP = re.compile(r'^\|[\s\-:|]+\|$')
RE_TABLE_ROW = re.compile(r'^\|(.+)\|$')
RE_HEADING = re.compile(r'^(#{1,6})\s+(.+)$')
RE_BULLET_LABEL = re.compile(r'^\s*[*\-+]\s+\*\*(.+?)[：:]\*\*\s*(.*)$')
RE_BULLET_PLAIN = re.compile(r'^\s*[*\-+]\s+(.+)$')
RE_NUMBERED = re.compile(r'^\s*(\d+)[.、）)]\s*(.+)$')
RE_META_LABEL = re.compile(r'^\*\*(.+?)[：:]\*\*\s*(.*)$')
RE_BOLD_ONLY = re.compile(r'^\*\*(.+?)\*\*\s*$')
RE_SECTION_BARE = re.compile(r'^([一二三四五六七八九十百千]+、)\s*(.+)$')
RE_BLOCKQUOTE = re.compile(r'^>\s*(.*)$')
RE_PLAIN_LABEL = re.compile(r'^([\u4e00-\u9fff]{2,8})[：:]\s*(.+)$')
RE_FENCE = re.compile(r'^`{3,}(\w*)$')
RE_TASK = re.compile(r'^(\s*)[*\-+]\s+\[([ xX])\]\s+(.+)$')
RE_IMAGE_LINE = re.compile(r'^!\[([^\]]*)\]\(([^)]+)\)\s*$')
RE_NESTED_BULLET = re.compile(r'^(\s{2,})[*\-+]\s+(.+)$')
RE_NESTED_NUMBERED = re.compile(r'^(\s{2,})(\d+)[.、）)]\s*(.+)$')


# ── YAML Front Matter ──

RE_FRONT_MATTER = re.compile(r'\A---\r?\n(.*?)\r?\n---\r?\n', re.DOTALL)

def parse_front_matter(text):
    """解析 YAML Front Matter，返回 (metadata_dict, remaining_text)"""
    m = RE_FRONT_MATTER.match(text)
    if not m:
        return {}, text
    try:
        meta = yaml.safe_load(m.group(1))
        if not isinstance(meta, dict):
            return {}, text
        return meta, text[m.end():]
    except yaml.YAMLError:
        return {}, text


# ── 内联文本处理 ──

RE_INLINE = re.compile(
    r'(`[^`]+`)'                          # 行内代码
    r'|(!\[([^\]]*)\]\(([^)]+)\))'        # 图片 ![alt](path)
    r'|(\[([^\]]+)\]\(([^)]+)\))'         # 超链接 [text](url)
    r'|(~~.+?~~)'                         # 删除线
    r'|(\*{2,3}.+?\*{2,3})'              # 加粗 / 加粗斜体
)


def strip_italic(text):
    text = re.sub(r'(?<!\*)\*([^*]+)\*(?!\*)', r'\1', text)
    text = re.sub(r'(?<!_)_([^_]+)_(?!_)', r'\1', text)
    return text


def parse_inline(text):
    """解析内联元素，返回结构化片段列表。

    每个片段是字典: {'type': 'text'|'code'|'link'|'image', ...}
    """
    if not text:
        return []

    result = []
    last_end = 0

    for m in RE_INLINE.finditer(text):
        start, end = m.span()
        if start > last_end:
            plain = strip_italic(text[last_end:start])
            if plain:
                result.append({'type': 'text', 'text': plain, 'bold': False})

        if m.group(1):
            result.append({'type': 'code', 'text': m.group(1)[1:-1]})
        elif m.group(2):
            result.append({'type': 'image', 'alt': m.group(3), 'path': m.group(4)})
        elif m.group(5):
            result.append({'type': 'link', 'text': m.group(6), 'url': m.group(7)})
        elif m.group(8):
            raw = m.group(8)
            result.append({'type': 'text', 'text': raw[2:-2], 'bold': False, 'strike': True})
        elif m.group(9):
            raw = m.group(9)
            if raw.startswith('***') and raw.endswith('***') and len(raw) > 6:
                result.append({'type': 'text', 'text': strip_italic(raw[3:-3]), 'bold': True})
            elif raw.startswith('**') and raw.endswith('**') and len(raw) > 4:
                result.append({'type': 'text', 'text': strip_italic(raw[2:-2]), 'bold': True})

        last_end = end

    if last_end < len(text):
        plain = strip_italic(text[last_end:])
        if plain:
            result.append({'type': 'text', 'text': plain, 'bold': False})

    return result


def parse_inline_bold(text):
    """兼容桥接：返回旧格式 (is_bold, text) 二元组列表"""
    parts = parse_inline(text)
    result = []
    for part in parts:
        if part['type'] == 'text':
            result.append((part.get('bold', False), part['text']))
        elif part['type'] == 'code':
            result.append((False, part['text']))
        elif part['type'] == 'link':
            result.append((False, part['text']))
        elif part['type'] == 'image':
            result.append((False, f'[图片: {part["alt"]}]'))
    return result


# ── Markdown 解析器（增强版）──

def parse_markdown(text):
    front_matter, text = parse_front_matter(text)
    lines = text.split('\n')
    blocks = []
    title_seen = bool(front_matter.get('title'))
    in_metadata = False
    table_headers = None
    table_rows = []
    in_code_block = False
    code_lang = ''
    code_lines = []

    def flush_table():
        nonlocal table_headers, table_rows
        if table_headers is not None:
            blocks.append({
                'type': 'table',
                'headers': table_headers,
                'rows': table_rows,
            })
            table_headers = None
            table_rows = []

    def flush_code_block():
        nonlocal in_code_block, code_lang, code_lines
        if in_code_block:
            blocks.append({
                'type': 'code_block',
                'language': code_lang,
                'code': '\n'.join(code_lines),
            })
            in_code_block = False
            code_lang = ''
            code_lines = []

    for line in lines:
        line = line.rstrip('\r')

        # 围栏代码块（最高优先级，代码块内不做任何其他解析）
        m = RE_FENCE.match(line.strip())
        if m:
            if in_code_block:
                flush_code_block()
            else:
                flush_table()
                in_code_block = True
                code_lang = m.group(1)
                code_lines = []
            continue

        if in_code_block:
            code_lines.append(line)
            continue

        if RE_TABLE_SEP.match(line):
            continue

        m = RE_TABLE_ROW.match(line)
        if m:
            cells = [c.strip() for c in m.group(1).split('|')]
            cells = [re.sub(r'\*\*(.+?)\*\*', r'\1', c) for c in cells]
            if table_headers is None:
                table_headers = cells
            else:
                table_rows.append(cells)
            continue

        flush_table()

        if RE_EMPTY.match(line):
            if blocks and blocks[-1]['type'] != 'empty':
                blocks.append({'type': 'empty'})
            continue

        if RE_HR.match(line):
            in_metadata = False
            if blocks and blocks[-1]['type'] != 'empty':
                blocks.append({'type': 'empty'})
            continue

        # 图片（独占一行）
        m = RE_IMAGE_LINE.match(line.strip())
        if m:
            in_metadata = False
            blocks.append({
                'type': 'image',
                'alt': m.group(1),
                'path': m.group(2),
            })
            continue

        m = RE_HEADING.match(line)
        if m:
            level = len(m.group(1))
            heading_text = m.group(2).strip()

            if not title_seen:
                blocks.append({'type': 'title', 'text': heading_text})
                title_seen = True
                in_metadata = True
                continue

            in_metadata = False
            if RE_SECTION_BARE.match(heading_text) or level <= 2:
                blocks.append({'type': 'section_header', 'text': heading_text})
            else:
                blocks.append({'type': 'sub_header', 'text': heading_text})
            continue

        m = RE_BLOCKQUOTE.match(line)
        if m:
            inner = m.group(1).strip()
            if inner:
                in_metadata = False
                blocks.append({'type': 'body', 'text': inner})
            continue

        # 任务列表（在普通列表之前检测）
        m = RE_TASK.match(line)
        if m:
            in_metadata = False
            blocks.append({
                'type': 'task_item',
                'text': m.group(3).strip(),
                'checked': m.group(2) != ' ',
            })
            continue

        # 嵌套列表（缩进 >= 2 个空格的列表项）
        m = RE_NESTED_BULLET.match(line)
        if m:
            in_metadata = False
            level = len(m.group(1)) // 2
            blocks.append({
                'type': 'list_item',
                'text': m.group(2).strip(),
                'level': level,
                'ordered': False,
            })
            continue

        m = RE_NESTED_NUMBERED.match(line)
        if m:
            in_metadata = False
            level = len(m.group(1)) // 2
            blocks.append({
                'type': 'list_item',
                'text': m.group(3).strip(),
                'level': level,
                'ordered': True,
                'number': m.group(2),
            })
            continue

        m = RE_BULLET_LABEL.match(line)
        if m:
            in_metadata = False
            blocks.append({
                'type': 'label_content',
                'label': m.group(1).strip(),
                'content': m.group(2).strip(),
            })
            continue

        m = RE_BULLET_PLAIN.match(line)
        if m:
            in_metadata = False
            inner = m.group(1).strip()
            m2 = RE_BOLD_ONLY.match(inner)
            if m2:
                blocks.append({'type': 'sub_header', 'text': m2.group(1).strip()})
            else:
                m3 = RE_META_LABEL.match(inner)
                if m3:
                    blocks.append({
                        'type': 'label_content',
                        'label': m3.group(1).strip(),
                        'content': m3.group(2).strip(),
                    })
                else:
                    blocks.append({'type': 'body', 'text': inner})
            continue

        m = RE_META_LABEL.match(line)
        if m:
            blocks.append({
                'type': 'label_content',
                'label': m.group(1).strip(),
                'content': m.group(2).strip(),
            })
            continue

        m = RE_BOLD_ONLY.match(line)
        if m:
            in_metadata = False
            blocks.append({'type': 'sub_header', 'text': m.group(1).strip()})
            continue

        m = RE_SECTION_BARE.match(line.strip())
        if m:
            in_metadata = False
            blocks.append({'type': 'section_header', 'text': line.strip()})
            continue

        m = RE_NUMBERED.match(line)
        if m:
            in_metadata = False
            num = m.group(1)
            rest = m.group(2).strip()

            m2 = RE_META_LABEL.match(rest)
            if m2:
                blocks.append({
                    'type': 'label_content',
                    'label': m2.group(1).strip(),
                    'content': m2.group(2).strip(),
                })
                continue

            m2 = RE_BOLD_ONLY.match(rest)
            if m2:
                blocks.append({'type': 'sub_header', 'text': f'{num}. {m2.group(1).strip()}'})
                continue

            if rest.startswith('**') and rest.endswith('**') and len(rest) > 4:
                blocks.append({'type': 'sub_header', 'text': f'{num}. {rest[2:-2].strip()}'})
                continue

            blocks.append({'type': 'sub_header', 'text': f'{num}. {rest}'})
            continue

        if in_metadata:
            m = RE_PLAIN_LABEL.match(line.strip())
            if m:
                blocks.append({
                    'type': 'label_content',
                    'label': m.group(1).strip(),
                    'content': m.group(2).strip(),
                })
                continue
            else:
                in_metadata = False

        stripped = line.strip()
        if stripped:
            if not title_seen and len(stripped) < 50:
                blocks.append({'type': 'title', 'text': stripped})
                title_seen = True
                in_metadata = True
            else:
                blocks.append({'type': 'body', 'text': stripped})

    flush_table()
    flush_code_block()

    fm_title = front_matter.get('title', '') if front_matter else ''
    has_title = any(b['type'] == 'title' for b in blocks)

    if fm_title:
        if has_title:
            for b in blocks:
                if b['type'] == 'title':
                    b['text'] = str(fm_title)
                    break
        else:
            blocks.insert(0, {'type': 'title', 'text': str(fm_title)})
    elif not has_title:
        blocks.insert(0, {
            'type': 'title',
            'text': f'会议纪要 - {datetime.now().strftime("%Y年%m月%d日")}',
        })

    if front_matter:
        FM_LABEL_MAP = {
            'date': '日期', 'author': '作者', 'department': '部门',
            'attendees': '参会人员', 'location': '地点',
            'subject': '主题', 'version': '版本',
        }
        title_idx = next((i for i, b in enumerate(blocks) if b['type'] == 'title'), 0)
        insert_pos = title_idx + 1
        for key, label in FM_LABEL_MAP.items():
            val = front_matter.get(key)
            if val is not None:
                blocks.insert(insert_pos, {
                    'type': 'label_content',
                    'label': label,
                    'content': str(val),
                })
                insert_pos += 1
        for key, val in front_matter.items():
            if key not in FM_LABEL_MAP and key != 'title' and val is not None:
                blocks.insert(insert_pos, {
                    'type': 'label_content',
                    'label': str(key),
                    'content': str(val),
                })
                insert_pos += 1

    return blocks
