# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ("C:\\Users\\Carlos\\AppData\\Local\\Programs\\Python\\Python314\\Lib\\site-packages\\customtkinter", "customtkinter"),
        ("C:\\Users\\Carlos\\AppData\\Local\\Programs\\Python\\Python314\\Lib\\site-packages\\matplotlib\\mpl-data", "matplotlib\\mpl-data"),
    ],
    hiddenimports=[
        "customtkinter",
        "matplotlib",
        "matplotlib.backends.backend_tkagg",
        "matplotlib.backends._backend_tk",
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
    a.binaries,
    a.datas,
    [],
    name='BibliotecaYugioh',
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
    icon='NONE',
)
