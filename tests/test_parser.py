"""parser 模块单元测试"""

import pytest
from converter.parser import (
    parse_markdown, parse_inline, parse_inline_bold, strip_italic,
)


class TestStripItalic:
    def test_single_star(self):
        assert strip_italic('hello *world* end') == 'hello world end'

    def test_single_underscore(self):
        assert strip_italic('hello _world_ end') == 'hello world end'

    def test_bold_not_stripped(self):
        assert strip_italic('hello **world** end') == 'hello **world** end'

    def test_no_italic(self):
        assert strip_italic('hello world') == 'hello world'


class TestParseInline:
    def test_plain_text(self):
        result = parse_inline('hello world')
        assert len(result) == 1
        assert result[0] == {'type': 'text', 'text': 'hello world', 'bold': False}

    def test_bold(self):
        result = parse_inline('hello **world** end')
        assert any(p['type'] == 'text' and p['bold'] and p['text'] == 'world' for p in result)

    def test_inline_code(self):
        result = parse_inline('use `print()` here')
        assert any(p['type'] == 'code' and p['text'] == 'print()' for p in result)

    def test_link(self):
        result = parse_inline('visit [Google](https://google.com) now')
        assert any(
            p['type'] == 'link' and p['text'] == 'Google' and p['url'] == 'https://google.com'
            for p in result
        )

    def test_image(self):
        result = parse_inline('![logo](img/logo.png)')
        assert any(p['type'] == 'image' and p['path'] == 'img/logo.png' for p in result)

    def test_empty(self):
        assert parse_inline('') == []
        assert parse_inline(None) == []

    def test_bold_italic(self):
        result = parse_inline('text ***important*** end')
        assert any(p['bold'] for p in result if p['type'] == 'text')


class TestParseInlineBold:
    def test_basic(self):
        result = parse_inline_bold('hello **world** end')
        assert len(result) == 3
        assert result[0] == (False, 'hello ')
        assert result[1] == (True, 'world')
        assert result[2] == (False, ' end')

    def test_code_not_bold(self):
        result = parse_inline_bold('use `code` here')
        assert any(not bold and text == 'code' for bold, text in result)


class TestParseMarkdown:
    def test_title(self):
        blocks = parse_markdown('# 会议纪要\n\n内容')
        assert blocks[0]['type'] == 'title'
        assert blocks[0]['text'] == '会议纪要'

    def test_auto_title_for_short_line(self):
        blocks = parse_markdown('项目周报')
        assert blocks[0]['type'] == 'title'
        assert blocks[0]['text'] == '项目周报'

    def test_auto_generated_title(self):
        blocks = parse_markdown('这是一段很长的正文内容，长度超过50个字符，所以不会被当作标题来处理，而是会被识别为正文内容。')
        assert blocks[0]['type'] == 'title'

    def test_section_header(self):
        blocks = parse_markdown('# 标题\n\n## 第一章')
        headers = [b for b in blocks if b['type'] == 'section_header']
        assert len(headers) == 1
        assert headers[0]['text'] == '第一章'

    def test_section_bare_chinese(self):
        blocks = parse_markdown('# 标题\n\n一、会议议题')
        headers = [b for b in blocks if b['type'] == 'section_header']
        assert len(headers) == 1
        assert headers[0]['text'] == '一、会议议题'

    def test_label_content(self):
        blocks = parse_markdown('# 标题\n\n**时间：** 2026年1月1日')
        labels = [b for b in blocks if b['type'] == 'label_content']
        assert len(labels) == 1
        assert labels[0]['label'] == '时间'
        assert labels[0]['content'] == '2026年1月1日'

    def test_bullet_label(self):
        blocks = parse_markdown('# 标题\n\n- **负责人：** 张三')
        labels = [b for b in blocks if b['type'] == 'label_content']
        assert len(labels) == 1
        assert labels[0]['label'] == '负责人'

    def test_body(self):
        blocks = parse_markdown('# 标题\n\n这是正文内容。')
        bodies = [b for b in blocks if b['type'] == 'body']
        assert len(bodies) == 1
        assert bodies[0]['text'] == '这是正文内容。'

    def test_table(self):
        md = '# 标题\n\n| 姓名 | 部门 |\n| --- | --- |\n| 张三 | 技术 |\n| 李四 | 产品 |'
        blocks = parse_markdown(md)
        tables = [b for b in blocks if b['type'] == 'table']
        assert len(tables) == 1
        assert tables[0]['headers'] == ['姓名', '部门']
        assert len(tables[0]['rows']) == 2

    def test_code_block(self):
        md = '# 标题\n\n```python\nprint("hello")\n```'
        blocks = parse_markdown(md)
        codes = [b for b in blocks if b['type'] == 'code_block']
        assert len(codes) == 1
        assert codes[0]['language'] == 'python'
        assert 'print("hello")' in codes[0]['code']

    def test_task_item(self):
        md = '# 标题\n\n- [x] 完成任务\n- [ ] 待做任务'
        blocks = parse_markdown(md)
        tasks = [b for b in blocks if b['type'] == 'task_item']
        assert len(tasks) == 2
        assert tasks[0]['checked'] is True
        assert tasks[1]['checked'] is False

    def test_image(self):
        md = '# 标题\n\n![截图](screenshots/img1.png)'
        blocks = parse_markdown(md)
        images = [b for b in blocks if b['type'] == 'image']
        assert len(images) == 1
        assert images[0]['path'] == 'screenshots/img1.png'

    def test_nested_list(self):
        md = '# 标题\n\n- 一级\n  - 二级'
        blocks = parse_markdown(md)
        items = [b for b in blocks if b['type'] in ('body', 'list_item')]
        nested = [b for b in blocks if b['type'] == 'list_item' and b.get('level', 0) > 0]
        assert len(nested) >= 1

    def test_blockquote(self):
        md = '# 标题\n\n> 引用内容'
        blocks = parse_markdown(md)
        bodies = [b for b in blocks if b['type'] == 'body']
        assert any('引用内容' in b['text'] for b in bodies)

    def test_empty_input(self):
        blocks = parse_markdown('')
        assert blocks[0]['type'] == 'title'

    def test_horizontal_rule(self):
        md = '# 标题\n\n---\n\n内容'
        blocks = parse_markdown(md)
        assert not any(b.get('text', '') == '---' for b in blocks)

    def test_numbered_list(self):
        md = '# 标题\n\n1. 第一项\n2. 第二项'
        blocks = parse_markdown(md)
        subs = [b for b in blocks if b['type'] == 'sub_header']
        assert len(subs) == 2

    def test_bold_only_line(self):
        md = '# 标题\n\n**重要提醒**'
        blocks = parse_markdown(md)
        subs = [b for b in blocks if b['type'] == 'sub_header']
        assert any('重要提醒' in s['text'] for s in subs)

    def test_sub_header(self):
        md = '# 标题\n\n### 三级标题'
        blocks = parse_markdown(md)
        subs = [b for b in blocks if b['type'] == 'sub_header']
        assert any('三级标题' in s['text'] for s in subs)
