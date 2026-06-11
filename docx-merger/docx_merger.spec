# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['docx_merger.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        'docx.shared',
        'PIL',
        'PIL.Image',
        'lxml',
        'lxml.etree',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='DOCX文件合并工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
