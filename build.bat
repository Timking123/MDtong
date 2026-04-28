@echo off
chcp 65001 >nul
echo ===== MD通 - 构建脚本 =====
echo.

echo [1/3] 安装依赖...
pip install -r requirements.txt
if errorlevel 1 (
    echo 依赖安装失败！
    pause
    exit /b 1
)

echo.
echo [2/3] 尝试安装可选拖拽支持...
pip install tkinterdnd2>=0.4.2 2>nul
if errorlevel 1 (
    echo 提示：tkinterdnd2 安装失败，拖拽功能将不可用（不影响其他功能）
)

echo.
echo [3/3] 打包为 exe...
pyinstaller --name "MD通" --windowed --noconfirm --add-data "templates;templates" --add-data "config.json;." --hidden-import anthropic --hidden-import anthropic._exceptions --hidden-import docx2pdf --hidden-import lxml --hidden-import lxml.etree --hidden-import PIL --hidden-import PIL.Image --collect-submodules anthropic --collect-submodules docx convert.py

if errorlevel 1 (
    echo 打包失败！
    pause
    exit /b 1
)

echo.
echo ===== 构建完成！=====
echo 输出目录: dist\MD通\
pause
