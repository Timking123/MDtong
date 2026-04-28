"""DOCX → Markdown 反向转换（增强版）"""

import re
from docx import Document
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH


def _is_center_aligned(paragraph):
    pPr = paragraph._element.find(qn('w:pPr'))
    if pPr is not None:
        jc = pPr.find(qn('w:jc'))
        if jc is not None and jc.get(qn('w:val')) == 'center':
            return True
    return paragraph.alignment is not None and paragraph.alignment == 1


def _is_bold_paragraph(paragraph):
    if not paragraph.runs:
        return False
    return all(r.bold for r in paragraph.runs if r.text.strip())


def _get_max_font_size(paragraph):
    sizes = []
    for run in paragraph.runs:
        if run.font.size:
            sizes.append(run.font.size)
    return max(sizes) if sizes else None


def _get_list_info(paragraph):
    """检测段落是否为列表项，返回 (level, num_id, is_ordered) 或 None"""
    pPr = paragraph._element.find(qn('w:pPr'))
    if pPr is None:
        return None
    numPr = pPr.find(qn('w:numPr'))
    if numPr is None:
        return None
    ilvl_el = numPr.find(qn('w:ilvl'))
    numId_el = numPr.find(qn('w:numId'))
    if ilvl_el is None or numId_el is None:
        return None
    level = int(ilvl_el.get(qn('w:val'), '0'))
    num_id = int(numId_el.get(qn('w:val'), '0'))
    if num_id == 0:
        return None
    return level, num_id, None


def _get_indent_level(paragraph):
    """通过段落缩进推断列表层级"""
    pPr = paragraph._element.find(qn('w:pPr'))
    if pPr is None:
        return 0
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        return 0
    left = ind.get(qn('w:left'), '0')
    try:
        twips = int(left)
        return max(0, twips // 720)
    except (ValueError, TypeError):
        return 0


def _has_code_font(paragraph):
    """检测段落是否使用等宽字体（代码块特征）"""
    code_fonts = {'consolas', 'courier', 'courier new', 'monospace', 'source code pro', 'fira code'}
    for run in paragraph.runs:
        if run.font.name and run.font.name.lower() in code_fonts:
            return True
    return False


def _has_shading(paragraph):
    """检测段落是否有底色（代码块特征）"""
    pPr = paragraph._element.find(qn('w:pPr'))
    if pPr is not None:
        shd = pPr.find(qn('w:shd'))
        if shd is not None:
            fill = shd.get(qn('w:fill'), '').upper()
            if fill and fill not in ('FFFFFF', 'AUTO', ''):
                return True
    return False


def _extract_hyperlinks(paragraph):
    """提取段落中的超链接"""
    links = {}
    for link_el in paragraph._element.findall(qn('w:hyperlink')):
        r_id = link_el.get(qn('r:id'))
        text_parts = []
        for run_el in link_el.findall(qn('w:r')):
            for t_el in run_el.findall(qn('w:t')):
                if t_el.text:
                    text_parts.append(t_el.text)
        link_text = ''.join(text_parts)
        if r_id and link_text:
            try:
                rel = paragraph.part.rels.get(r_id)
                if rel and hasattr(rel, 'target_ref'):
                    links[link_text] = rel.target_ref
            except Exception:
                pass
    return links


def _runs_to_markdown(paragraph):
    """将段落的 runs 转换为 Markdown 文本，支持加粗、行内代码、超链接"""
    links = _extract_hyperlinks(paragraph)
    code_fonts = {'consolas', 'courier', 'courier new', 'monospace', 'source code pro', 'fira code'}

    parts = []
    link_el_texts = set()
    for link_el in paragraph._element.findall(qn('w:hyperlink')):
        for run_el in link_el.findall(qn('w:r')):
            for t_el in run_el.findall(qn('w:t')):
                if t_el.text:
                    link_el_texts.add(t_el.text)

    for run in paragraph.runs:
        text = run.text
        if not text:
            continue

        is_code = run.font.name and run.font.name.lower() in code_fonts
        has_shd = False
        rPr = run._element.find(qn('w:rPr'))
        if rPr is not None:
            shd = rPr.find(qn('w:shd'))
            if shd is not None:
                fill = shd.get(qn('w:fill'), '').upper()
                if fill and fill not in ('FFFFFF', 'AUTO', ''):
                    has_shd = True

        if text in links:
            parts.append(f'[{text}]({links[text]})')
        elif is_code or has_shd:
            parts.append(f'`{text}`')
        elif run.bold:
            parts.append(f'**{text}**')
        else:
            parts.append(text)

    result = ''.join(parts)
    while '****' in result:
        result = result.replace('****', '')
    return result


def _detect_bullet_char(text):
    """检测文本是否以项目符号开头"""
    bullets = ('•', '·', '◦', '▪', '▸', '►', '●', '○', '■', '□', '☑', '☐', '✓', '✔', '✗', '✘')
    stripped = text.lstrip()
    for b in bullets:
        if stripped.startswith(b):
            return b
    return None


def _detect_task_item(text):
    """检测文本是否为任务列表项，返回 (checked, content) 或 None"""
    stripped = text.lstrip()
    if stripped.startswith('☑ ') or stripped.startswith('☑'):
        return True, stripped[1:].lstrip()
    if stripped.startswith('☐ ') or stripped.startswith('☐'):
        return False, stripped[1:].lstrip()
    if stripped.startswith('✓ ') or stripped.startswith('✔ '):
        return True, stripped[1:].lstrip()
    if stripped.startswith('✗ ') or stripped.startswith('✘ '):
        return False, stripped[1:].lstrip()
    return None


def _detect_numbered_item(text):
    """检测文本是否为编号列表项"""
    m = re.match(r'^(\d+)[.、）)]\s+(.+)$', text.lstrip())
    if m:
        return m.group(1), m.group(2)
    return None


def _table_to_markdown(table):
    lines = []
    rows = table.rows
    if not rows:
        return ''

    header_cells = [cell.text.strip() for cell in rows[0].cells]
    lines.append('| ' + ' | '.join(header_cells) + ' |')
    lines.append('| ' + ' | '.join(['---'] * len(header_cells)) + ' |')

    for row in rows[1:]:
        cells = [cell.text.strip() for cell in row.cells]
        while len(cells) < len(header_cells):
            cells.append('')
        lines.append('| ' + ' | '.join(cells[:len(header_cells)]) + ' |')

    return '\n'.join(lines)


def _extract_images(paragraph, doc):
    """提取段落中的图片信息"""
    images = []
    for run in paragraph.runs:
        drawing_els = run._element.findall(qn('w:drawing'))
        for drawing in drawing_els:
            for inline in drawing.findall('.//' + qn('wp:inline')):
                doc_pr = inline.find(qn('wp:docPr'))
                alt = ''
                if doc_pr is not None:
                    alt = doc_pr.get('descr', '') or doc_pr.get('name', '')
                images.append(alt)
            for anchor in drawing.findall('.//' + qn('wp:anchor')):
                doc_pr = anchor.find(qn('wp:docPr'))
                alt = ''
                if doc_pr is not None:
                    alt = doc_pr.get('descr', '') or doc_pr.get('name', '')
                images.append(alt)
    return images


def docx_to_markdown(docx_path):
    """将 .docx 文件转换为 Markdown 文本"""
    doc = Document(docx_path)
    md_lines = []
    prev_empty = False
    in_code_block = False
    code_lines = []
    ordered_counters = {}

    body = doc.element.body
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'p':
            from docx.text.paragraph import Paragraph
            para = Paragraph(child, doc)
            text = para.text.strip()

            images = _extract_images(para, doc)
            if images:
                if in_code_block:
                    md_lines.append('```')
                    in_code_block = False
                    code_lines = []
                for alt in images:
                    md_lines.append(f'![{alt}](image)')
                    prev_empty = False
                if not text:
                    continue

            is_code_line = _has_code_font(para) or _has_shading(para)

            if is_code_line and text:
                if not in_code_block:
                    if not prev_empty and md_lines:
                        md_lines.append('')
                    md_lines.append('```')
                    in_code_block = True
                code_lines.append(para.text.rstrip())
                md_lines.append(para.text.rstrip())
                prev_empty = False
                continue
            elif in_code_block:
                md_lines.append('```')
                md_lines.append('')
                in_code_block = False
                code_lines = []

            if not text:
                if not prev_empty:
                    md_lines.append('')
                    prev_empty = True
                continue

            prev_empty = False
            font_size = _get_max_font_size(para)

            list_info = _get_list_info(para)
            if list_info:
                level, num_id, _ = list_info
                indent = '  ' * level
                md_text = _runs_to_markdown(para)
                key = (num_id, level)
                if key not in ordered_counters:
                    ordered_counters[key] = 1
                else:
                    ordered_counters[key] += 1
                md_lines.append(f'{indent}- {md_text}')
                continue

            task = _detect_task_item(text)
            if task:
                checked, content = task
                marker = 'x' if checked else ' '
                md_lines.append(f'- [{marker}] {content}')
                continue

            bullet = _detect_bullet_char(text)
            if bullet:
                indent_level = _get_indent_level(para)
                indent = '  ' * indent_level
                content = text.lstrip()[len(bullet):].lstrip()
                md_text = _runs_to_markdown(para)
                if bullet in md_text:
                    md_text = md_text[md_text.index(bullet) + len(bullet):].lstrip()
                else:
                    md_text = content
                md_lines.append(f'{indent}- {md_text}')
                continue

            numbered = _detect_numbered_item(text)
            if numbered and not _is_center_aligned(para):
                num, content = numbered
                indent_level = _get_indent_level(para)
                indent = '  ' * indent_level
                md_lines.append(f'{indent}{num}. {content}')
                continue

            if _is_center_aligned(para) and font_size and font_size >= 152400:
                md_lines.append(f'# {text}')
                continue

            if _is_bold_paragraph(para):
                if font_size and font_size >= 140000:
                    md_lines.append(f'## {text}')
                else:
                    md_lines.append(f'**{text}**')
                continue

            md_text = _runs_to_markdown(para)
            md_lines.append(md_text)

        elif tag == 'tbl':
            if in_code_block:
                md_lines.append('```')
                in_code_block = False
                code_lines = []
            from docx.table import Table
            tbl = Table(child, doc)
            prev_empty = False
            md_lines.append('')
            md_lines.append(_table_to_markdown(tbl))
            md_lines.append('')

    if in_code_block:
        md_lines.append('```')

    result = '\n'.join(md_lines)
    while '\n\n\n' in result:
        result = result.replace('\n\n\n', '\n\n')
    return result.strip() + '\n'
