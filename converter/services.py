"""转换服务层。

统一封装文本读取、DOCX 导入、DOCX/PDF/HTML 导出等流程，供 CLI 和 GUI 复用。
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Callable, Optional

from .generator import SCRIPT_DIR, convert as convert_markdown_to_docx
from .html_export import export_html
from .reverse import docx_to_markdown
from .validator import has_errors, validate_conversion_input

ENCODINGS = ('utf-8-sig', 'utf-8', 'gbk')
ProgressCallback = Callable[[int, str], None]


class ConversionCancelled(InterruptedError):
    """用户主动取消转换任务。"""


@dataclass(slots=True)
class ConversionRequest:
    """单次转换请求。"""

    text: str
    output_path: Optional[str] = None
    output_dir: Optional[str] = None
    input_path: Optional[str] = None
    doc_features: Optional[dict] = None
    format_settings: Optional[dict] = None
    base_dir: Optional[str] = None
    also_pdf: bool = False
    also_html: bool = False


@dataclass(slots=True)
class ConversionResult:
    """单次转换结果。"""

    docx_path: str
    pdf_path: Optional[str] = None
    html_path: Optional[str] = None
    warnings: list[str] = field(default_factory=list)


def ensure_not_cancelled(cancel_event=None) -> None:
    """如果取消事件已触发则抛出统一取消异常。"""
    if cancel_event is not None and cancel_event.is_set():
        raise ConversionCancelled('用户取消')


def read_text_file(path: str, encodings: tuple[str, ...] = ENCODINGS) -> str:
    """按常见中文编码读取文本文件。"""
    last_error: Optional[Exception] = None
    for enc in encodings:
        try:
            with open(path, 'r', encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
    raise ValueError(f'无法读取文件，编码不支持：{path}') from last_error


def load_input_as_markdown(path: str) -> str:
    """读取 Markdown/TXT 或将 DOCX 反向转换为 Markdown。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == '.docx':
        return docx_to_markdown(path)
    return read_text_file(path)


def default_output_path(input_path: Optional[str], output_dir: Optional[str], suffix: str) -> str:
    """基于输入文件和输出目录生成默认输出路径。"""
    out_dir = output_dir or SCRIPT_DIR
    if input_path:
        basename = os.path.splitext(os.path.basename(input_path))[0]
    else:
        basename = 'MD通文档'
    return os.path.join(out_dir, f'{basename}{suffix}')


def convert_request(
    request: ConversionRequest,
    progress_cb: Optional[ProgressCallback] = None,
    cancel_event=None,
) -> ConversionResult:
    """执行 Markdown 到 DOCX 的统一转换流程。"""
    ensure_not_cancelled(cancel_event)
    output_dir = request.output_dir or (os.path.dirname(request.output_path) if request.output_path else SCRIPT_DIR)
    if output_dir and not os.path.isdir(output_dir):
        raise FileNotFoundError(f'输出目录不存在：{output_dir}')

    issues = validate_conversion_input(
        request.text,
        base_dir=request.base_dir,
        output_path=request.output_path,
        output_dir=output_dir,
    )
    blocking = [issue.format() for issue in issues if issue.level == 'error']
    if has_errors(issues):
        raise ValueError('转换前检查未通过：\n' + '\n'.join(blocking))

    ensure_not_cancelled(cancel_event)
    docx_path = convert_markdown_to_docx(
        request.text,
        request.output_path,
        progress_cb=progress_cb,
        doc_features=request.doc_features,
        base_dir=request.base_dir,
        format_settings=request.format_settings,
        output_dir=output_dir,
        cancel_event=cancel_event,
    )

    result = ConversionResult(docx_path=docx_path, warnings=[issue.format() for issue in issues if issue.level != 'error'])

    ensure_not_cancelled(cancel_event)
    if request.also_pdf:
        from .pdf_export import export_pdf

        result.pdf_path = export_pdf(docx_path)

    ensure_not_cancelled(cancel_event)
    if request.also_html:
        html_path = os.path.splitext(docx_path)[0] + '.html'
        result.html_path = export_html(request.text, html_path)

    return result


def convert_file(
    input_path: str,
    output_path: Optional[str] = None,
    output_dir: Optional[str] = None,
    progress_cb: Optional[ProgressCallback] = None,
    doc_features: Optional[dict] = None,
    format_settings: Optional[dict] = None,
    also_pdf: bool = False,
    cancel_event=None,
) -> ConversionResult:
    """读取输入文件并转换为 DOCX。"""
    ensure_not_cancelled(cancel_event)
    text = load_input_as_markdown(input_path)
    base_dir = os.path.dirname(os.path.abspath(input_path))
    request = ConversionRequest(
        text=text,
        output_path=output_path,
        output_dir=output_dir,
        input_path=input_path,
        doc_features=doc_features,
        format_settings=format_settings,
        base_dir=base_dir,
        also_pdf=also_pdf,
    )
    return convert_request(request, progress_cb=progress_cb, cancel_event=cancel_event)
