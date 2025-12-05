# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['pdf_editor_v3.1_멀티창_보안기능변경.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('data/kunhwa_logo.png', 'data'),
        ('data/kunhwa_logo.ico', 'data'),
        ('사용자_설명서.txt', '.'),
        ('맥어드레스.xlsx', '.'),
        ('users.json.enc', '.')
    ],
    hiddenimports=[
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.ttk',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'fitz',
        'fitz.fitz',
        'subprocess',
        'os',
        'sys',
        'json',
        'base64',
        'hashlib',
        'uuid',
        're',
        'tempfile',
        'time',
        'functools'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Kunhwa_PDF_Editor_v3.1_배포용',
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
    icon='data/kunhwa_logo.ico',
    version_file=None,
    uac_admin=False,
    uac_uiaccess=False,
)
