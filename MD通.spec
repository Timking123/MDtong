# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['anthropic', 'anthropic._exceptions', 'docx2pdf', 'lxml', 'lxml.etree', 'PIL', 'PIL.Image', 'tkinterdnd2', 'markdown', 'markdown.extensions.tables', 'markdown.extensions.fenced_code', 'markdown.extensions.toc']
hiddenimports += collect_submodules('anthropic')
hiddenimports += collect_submodules('docx')


a = Analysis(
    ['convert.py'],
    pathex=[],
    binaries=[],
    datas=[('templates', 'templates'), ('config.json', '.')],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='MD通',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='MD通',
)
