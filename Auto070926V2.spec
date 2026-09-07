# -*- mode: python ; coding: utf-8 -*-


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
        'bs4',
        'lxml',
        'lxml.etree',
        'win32com',
        'win32com.client',
        'pythoncom',
        'matplotlib',
        'matplotlib.pyplot',
        'numpy',
        'curl_cffi',
        'PIL',
        'PIL.Image',
    ],
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
    name='Auto070926V2',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
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
    name='Auto070926V2',
)
