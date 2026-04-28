# MD通

MD通是一款面向中文办公场景的 Markdown 文档转换工具，支持将 Markdown 转换为 DOCX，并提供模板、格式校验、批量转换、DOCX 反向解析、HTML/PDF 导出和 AI 润色相关能力。

## 版本

当前版本：3.0.0

## 功能特性

- Markdown 转 DOCX
- DOCX 转 Markdown
- 多模板支持
- 文档格式校验
- 批量转换
- HTML/PDF 导出能力
- 图形界面操作
- 可选 AI 润色配置

## 快速开始

### 直接运行

下载发布包后运行 `MD通.exe`。

### 从源码运行

```bash
python -m pip install -r requirements.txt
python convert.py
```

### 打包 EXE

```bash
python -m pip install pyinstaller>=6.0.0
python -m PyInstaller --clean --noconfirm MD通.spec
```

打包结果位于 `dist/MD通/`。

## 配置说明

- `config.json` 保存本地配置，默认不包含密钥。
- `.env.example` 提供环境变量示例。
- 请不要将真实 API Key、私密路径或个人数据提交到公开仓库。

## 开源许可

本项目使用 MIT License 开源。
