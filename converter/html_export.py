"""HTML 导出 — 将 Markdown 文本转换为 HTML 文件"""

import re
import html
import markdown


_RE_DANGEROUS_TAGS = re.compile(
    r'<\s*/?\s*(script|style|iframe|object|embed|form|input|textarea|button|link|meta|base)\b[^>]*>',
    re.IGNORECASE,
)
_RE_EVENT_HANDLER = re.compile(r'\s+on\w+\s*=\s*["\'][^"\']*["\']', re.IGNORECASE)
_RE_JS_URI = re.compile(r'(href|src|action)\s*=\s*["\']?\s*javascript:', re.IGNORECASE)


def _sanitize_html(raw_html):
    """移除 HTML 中的危险标签、事件处理器和 javascript: URI"""
    result = _RE_DANGEROUS_TAGS.sub('', raw_html)
    result = _RE_EVENT_HANDLER.sub('', result)
    result = _RE_JS_URI.sub(r'\1="', result)
    return result


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
body {{
    font-family: 'Microsoft YaHei', sans-serif;
    max-width: 800px;
    margin: 40px auto;
    padding: 0 20px;
    line-height: 1.6;
    color: #333;
}}
h1 {{ text-align: center; color: #1a56db; }}
h2 {{ color: #1a56db; border-bottom: 1px solid #eee; padding-bottom: 5px; }}
h3 {{ color: #333; }}
table {{
    border-collapse: collapse;
    width: 100%;
    margin: 15px 0;
}}
th, td {{
    border: 1px solid #ddd;
    padding: 8px 12px;
    text-align: left;
}}
th {{ background-color: #f5f5f5; font-weight: bold; }}
code {{
    background: #f5f5f5;
    padding: 2px 5px;
    border-radius: 3px;
    font-family: Consolas, monospace;
}}
pre {{
    background: #f5f5f5;
    padding: 15px;
    border-radius: 5px;
    overflow-x: auto;
}}
blockquote {{
    border-left: 4px solid #ddd;
    padding-left: 15px;
    color: #666;
    margin: 10px 0;
}}
img {{ max-width: 100%; }}
a {{ color: #0563C1; }}
</style>
</head>
<body>
{content}
</body>
</html>"""


def export_html(md_text, output_path):
    """将 Markdown 文本转换为 HTML 文件。

    Args:
        md_text: Markdown 原文
        output_path: 输出 .html 文件路径
    Returns:
        输出文件路径
    """
    html_content = markdown.markdown(
        md_text,
        extensions=['tables', 'fenced_code', 'toc'],
    )

    lines = md_text.strip().split('\n')
    title = 'MD通文档'
    for line in lines:
        line = line.strip()
        if line.startswith('# '):
            title = line[2:].strip()
            break
        if line and not line.startswith('#'):
            title = line[:50]
            break

    html_content = _sanitize_html(html_content)
    full_html = HTML_TEMPLATE.format(title=html.escape(title), content=html_content)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(full_html)

    return output_path
