"""validator 模块单元测试"""

import os
import tempfile

from converter.validator import has_errors, validate_conversion_input, validate_markdown, validate_output_path


class TestValidateMarkdown:
    def test_empty_content_is_error(self):
        issues = validate_markdown('')
        assert has_errors(issues)
        assert issues[0].code == 'empty_content'

    def test_missing_image_is_warning(self):
        issues = validate_markdown('# 标题\n\n![图](missing.png)', base_dir=tempfile.gettempdir())
        assert any(i.code == 'image_missing' and i.level == 'warning' for i in issues)

    def test_heading_skip_is_warning(self):
        issues = validate_markdown('# H1\n\n### H3')
        assert any(i.code == 'heading_skip' for i in issues)

    def test_table_separator_warning(self):
        issues = validate_markdown('# 标题\n\n| A | B |\n正文')
        assert any(i.code == 'table_separator' for i in issues)

    def test_remote_image_is_ignored(self):
        issues = validate_markdown('# 标题\n\n![图](https://example.com/a.png)')
        assert not any(i.code == 'image_missing' for i in issues)


class TestValidateOutputPath:
    def test_missing_output_dir_is_error(self):
        missing = os.path.join(tempfile.gettempdir(), 'mdtong_missing_dir_xxx')
        issues = validate_output_path(output_dir=missing)
        assert any(i.code == 'output_dir_missing' for i in issues)

    def test_valid_output_dir(self):
        issues = validate_output_path(output_dir=tempfile.gettempdir())
        assert not has_errors(issues)


class TestValidateConversionInput:
    def test_combines_markdown_and_output_checks(self):
        missing = os.path.join(tempfile.gettempdir(), 'mdtong_missing_dir_xxx')
        issues = validate_conversion_input('# 标题\n\n内容', output_dir=missing)
        assert any(i.code == 'output_dir_missing' for i in issues)
