# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\Rodrigo Peixoto\\Documents\\GitHub\\IFDIGITAL_3.0\\src\\icone ifdigital.ico', 'src'), ('C:\\Users\\Rodrigo Peixoto\\Documents\\GitHub\\IFDIGITAL_3.0\\src\\01florest.png', 'src')],
    hiddenimports=['openpyxl', 'pandas', 'tkinter', 'configparser', 'numpy', 'xlsxwriter', 'PIL.Image', 'PIL.ImageTk', 'tkinter.PhotoImage'],
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
    name='main',
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
    icon=['C:\\Users\\Rodrigo Peixoto\\Downloads\\icone ifdigital.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='main',
)
