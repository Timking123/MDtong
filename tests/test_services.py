"""services 模块单元测试"""

import os
import tempfile
import threading

import pytest

from converter.services import ConversionRequest, convert_file, convert_request, read_text_file


class TestReadTextFile:
    def test_read_utf8(self):
        with tempfile.NamedTemporaryFile('w', suffix='.md', delete=False, encoding='utf-8') as f:
            f.write('# 标题\n\n内容')
            path = f.name
        try:
            assert '标题' in read_text_file(path)
        finally:
            os.unlink(path)


class TestConvertFile:
    def test_convert_file_to_docx(self):
        with tempfile.NamedTemporaryFile('w', suffix='.md', delete=False, encoding='utf-8') as f:
            f.write('# 服务测试\n\n内容')
            input_path = f.name
        output_path = tempfile.NamedTemporaryFile(suffix='.docx', delete=False).name
        os.unlink(output_path)
        try:
            result = convert_file(input_path, output_path=output_path)
            assert os.path.isfile(result.docx_path)
        finally:
            if os.path.exists(input_path):
                os.unlink(input_path)
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_cancel_before_convert(self):
        event = threading.Event()
        event.set()
        request = ConversionRequest(text='# 取消\n\n内容', output_dir=tempfile.gettempdir())
        with pytest.raises(InterruptedError):
            convert_request(request, cancel_event=event)

    def test_validation_error_blocks_missing_output_dir(self):
        missing_dir = os.path.join(tempfile.gettempdir(), 'mdtong_missing_dir_yyy')
        request = ConversionRequest(text='# 标题\n\n内容', output_dir=missing_dir)
        with pytest.raises(FileNotFoundError):
            convert_request(request)
