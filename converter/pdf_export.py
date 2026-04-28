"""PDF 导出 — 将 DOCX 文件转换为 PDF"""

import os
from docx2pdf import convert as docx2pdf_convert


def export_pdf(docx_path, pdf_path=None):
    """
    将 .docx 转换为 .pdf。

    Args:
        docx_path: 输入 .docx 文件路径
        pdf_path: 输出 .pdf 文件路径（默认同目录同名）
    Returns:
        输出的 PDF 文件路径
    """
    if not os.path.isfile(docx_path):
        raise FileNotFoundError(f'找不到文件: {docx_path}')

    if pdf_path is None:
        pdf_path = os.path.splitext(docx_path)[0] + '.pdf'

    docx2pdf_convert(docx_path, pdf_path)

    if not os.path.isfile(pdf_path):
        raise RuntimeError('PDF 导出失败，请确认已安装 Microsoft Word')

    return pdf_path
