# -*- mode: python ; coding: utf-8 -*-

import os
import sys

_tcl_root = os.path.join(sys.base_prefix, 'tcl')
_tcl_dir = os.path.join(_tcl_root, 'tcl8.6')
_tk_dir = os.path.join(_tcl_root, 'tk8.6')

extra_datas = [
    ('.\\data\\kunhwa_logo.png', 'data'),
    ('.\\users.json.enc', '.'),
]

if os.path.isdir(_tcl_dir):
    extra_datas.append((_tcl_dir, 'tcl/tcl8.6'))
if os.path.isdir(_tk_dir):
    extra_datas.append((_tk_dir, 'tcl/tk8.6'))

a = Analysis(
    ['pdf_editor_v3.1_멀티창_보안기능변경_2.py'],
    pathex=[],
    binaries=[],
    datas=extra_datas,
    hiddenimports=[],
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
    a.binaries,
    a.datas,
    [],
    name='Kunhwa_PDF_Editor_v3.1',
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
    icon=['data\\kunhwa_logo.ico'],
)
