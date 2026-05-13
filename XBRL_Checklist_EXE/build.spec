# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 빌드 스펙 — XBRL CoE Checklist 1-1
#
# 빌드 방법 (XBRL_Checklist_EXE 폴더에서 실행):
#   pip install pyinstaller pandas openpyxl
#   pyinstaller build.spec
#
# 결과물: dist/XBRL_Checklist_1_1/XBRL_Checklist_1_1  (macOS/Linux)
#          dist/XBRL_Checklist_1_1/XBRL_Checklist_1_1.exe  (Windows)

import os

block_cipher = None

# 내장 데이터 파일 (소스경로, 번들 내 상대경로)
datas = [
    ('Axis_Domain_Check.xlsx',                   '.'),
    ('template/XBRL_CoE_Checklist_Result.xlsx',  'template'),
]

a = Analysis(
    ['main.py'],
    pathex=[os.path.abspath('.')],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'pandas',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'et_xmlfile',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['flask', 'jinja2', 'werkzeug'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='XBRL_Checklist_1_1',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # GUI 모드 (콘솔 창 없음)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='XBRL_Checklist_1_1',
)
