# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('UpdateDelek', 'UpdateDelek'),
        ('BituahRechev', 'BituahRechev'),
        ('Madadim', 'Madadim'),
        ('config.py', '.'),
    ],
    hiddenimports=[
        'tkinter',
        'tkinter.ttk',
        'selenium',
        'webdriver_manager',
        'webdriver_manager.chrome',
        'webdriver_manager.core',
        'requests',
        'beautifulsoup4',
        'bs4',
        'lxml',
        'lxml.etree',
        'pywin32',
        'win32com',
        'win32com.client',
        'matplotlib',
        'numpy',
        'PIL',
        'PIL.Image',
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
    [],
    exclude_binaries=True,
    name='UpdateDelek',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # ללא חלון קונסול - רק GUI
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # אפשר להוסיף אייקון כאן אם יש
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='UpdateDelek',
)

