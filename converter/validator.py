"""转换前文档检查器。"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Literal, Optional

from .parser import RE_TABLE_ROW, RE_TABLE_SEP, parse_front_matter

IssueLevel = Literal['error', 'warning', 'info']


@dataclass(slots=True)
class ValidationIssue:
    """文档检查结果项。"""

    level: IssueLevel
    message: str
    line: Optional[int] = None
    code: str = ''

    def format(self) -> str:
        prefix = {'error': '错误', 'warning': '警告', 'info': '提示'}.get(self.level, self.level)
        location = f'第 {self.line} 行：' if self.line else ''
        return f'[{prefix}] {location}{self.message}'


RE_HEADING = re.compile(r'^(#{1,6})\s+(.+)$')
RE_IMAGE = re.compile(r'!\[[^\]]*\]\(([^)]+)\)')
RE_FRONT_MATTER_RAW = re.compile(r'\A---\r?\n(.*?)\r?\n---\r?\n', re.DOTALL)


def _is_remote_or_anchor(path: str) -> bool:
    lowered = path.lower().strip()
    return lowered.startswith(('http://', 'https://', 'data:', '#', 'mailto:'))


def _line_number_for_offset(text: str, offset: int) -> int:
    return text[:offset].count('\n') + 1


def validate_markdown(text: str, base_dir: Optional[str] = None) -> list[ValidationIssue]:
    """检查 Markdown 内容中的常见问题。"""
    issues: list[ValidationIssue] = []
    if not text or not text.strip():
        return [ValidationIssue('error', '内容为空，无法转换', code='empty_content')]

    if text.startswith('---'):
        meta, _ = parse_front_matter(text)
        if not meta and RE_FRONT_MATTER_RAW.match(text):
            issues.append(ValidationIssue('warning', 'Front Matter 未解析为有效 YAML 对象', line=1, code='front_matter_invalid'))

    previous_level = 0
    in_code = False
    table_start_line: Optional[int] = None
    table_expected_cols: Optional[int] = None
    table_has_sep = False

    lines = text.splitlines()
    for idx, line in enumerate(lines, start=1):
        stripped = line.strip()
        if stripped.startswith('```'):
            in_code = not in_code
            continue
        if in_code:
            continue

        heading = RE_HEADING.match(line)
        if heading:
            level = len(heading.group(1))
            if previous_level and level > previous_level + 1:
                issues.append(ValidationIssue('warning', f'标题层级从 H{previous_level} 跳到 H{level}', idx, 'heading_skip'))
            previous_level = level

        row_match = RE_TABLE_ROW.match(line)
        if row_match:
            cols = len([c for c in row_match.group(1).split('|')])
            if table_expected_cols is None:
                table_expected_cols = cols
                table_start_line = idx
                table_has_sep = False
            elif cols != table_expected_cols:
                issues.append(ValidationIssue('warning', f'表格列数不一致，应为 {table_expected_cols} 列，实际 {cols} 列', idx, 'table_columns'))
            continue

        if table_expected_cols is not None and RE_TABLE_SEP.match(line):
            table_has_sep = True
            continue

        if table_expected_cols is not None:
            if not table_has_sep and table_start_line:
                issues.append(ValidationIssue('warning', '表格缺少分隔行，例如 | --- | --- |', table_start_line, 'table_separator'))
            table_expected_cols = None
            table_start_line = None
            table_has_sep = False

    if in_code:
        issues.append(ValidationIssue('warning', '代码块未闭合', len(lines), 'code_fence_unclosed'))
    if table_expected_cols is not None and not table_has_sep and table_start_line:
        issues.append(ValidationIssue('warning', '表格缺少分隔行，例如 | --- | --- |', table_start_line, 'table_separator'))

    root = base_dir or os.getcwd()
    for match in RE_IMAGE.finditer(text):
        img_path = match.group(1).strip().strip('"\'')
        if _is_remote_or_anchor(img_path):
            continue
        full_path = img_path if os.path.isabs(img_path) else os.path.join(root, img_path)
        if not os.path.isfile(full_path):
            issues.append(ValidationIssue('warning', f'图片文件不存在：{img_path}', _line_number_for_offset(text, match.start()), 'image_missing'))

    return issues


def validate_output_path(output_path: Optional[str] = None, output_dir: Optional[str] = None) -> list[ValidationIssue]:
    """检查输出路径和目录。"""
    issues: list[ValidationIssue] = []
    target_dir = output_dir or (os.path.dirname(output_path) if output_path else '')
    if target_dir and not os.path.isdir(target_dir):
        issues.append(ValidationIssue('error', f'输出目录不存在：{target_dir}', code='output_dir_missing'))
    if output_path and os.path.exists(output_path):
        try:
            with open(output_path, 'a+b'):
                pass
        except OSError:
            issues.append(ValidationIssue('error', f'输出文件可能被占用：{output_path}', code='output_file_locked'))
    return issues


def validate_conversion_input(text: str, base_dir: Optional[str] = None, output_path: Optional[str] = None, output_dir: Optional[str] = None) -> list[ValidationIssue]:
    """执行完整的转换前检查。"""
    return validate_markdown(text, base_dir=base_dir) + validate_output_path(output_path=output_path, output_dir=output_dir)


def has_errors(issues: list[ValidationIssue]) -> bool:
    """是否存在阻断型错误。"""
    return any(issue.level == 'error' for issue in issues)


def format_issues(issues: list[ValidationIssue], limit: int = 8) -> str:
    """格式化检查结果，便于 GUI/CLI 展示。"""
    shown = issues[:limit]
    text = '\n'.join(issue.format() for issue in shown)
    if len(issues) > limit:
        text += f'\n... 还有 {len(issues) - limit} 项'
    return text
