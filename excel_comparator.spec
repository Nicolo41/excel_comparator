# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['excel_comparator.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\nicol\\AppData\\Roaming\\Python\\Python311\\site-packages', '.'), ('C:\\\\Users\\\\nicol\\\\Documents\\\\Documents\\\\Code\\\\Python\\\\Vidanges\\\\ecarts_vidanges\\\\img\\\\logo_jr2.png', '.'), ('C:\\\\Users\\\\nicol\\\\Documents\\\\Documents\\\\Code\\\\Python\\\\Vidanges\\\\ecarts_vidanges\\\\img\\\\excel2.png', '.'), ('C:\\\\Users\\\\nicol\\\\Documents\\\\Documents\\\\Code\\\\Python\\\\Vidanges\\\\ecarts_vidanges\\\\img\\\\compare2.png', '.'), ('C:\\\\Users\\\\nicol\\\\Documents\\\\Documents\\\\Code\\\\Python\\\\Vidanges\\\\ecarts_vidanges\\\\img\\\\exit2.png', '.'), ('C:\\\\Users\\\\nicol\\\\Documents\\\\Documents\\\\Code\\\\Python\\\\Vidanges\\\\ecarts_vidanges\\\\img\\\\fichier2.png', '.'), ('C:\\\\Users\\\\nicol\\\\Documents\\\\Documents\\\\Code\\\\Python\\\\Vidanges\\\\ecarts_vidanges\\\\img\\\\folder2.png', '.'), ('C:\\\\Users\\\\nicol\\\\Documents\\\\Documents\\\\Code\\\\Python\\\\Vidanges\\\\ecarts_vidanges\\\\img\\\\error2.png', '.'), ('C:\\\\Users\\\\nicol\\\\Documents\\\\Documents\\\\Code\\\\Python\\\\Vidanges\\\\ecarts_vidanges\\\\img\\\\git2.png', '.'), ('C:\\\\Users\\\\nicol\\\\Documents\\\\Documents\\\\Code\\\\Python\\\\Vidanges\\\\ecarts_vidanges\\\\img\\\\ok.png', '.'), ('C:\\\\Users\\\\nicol\\\\Documents\\\\Documents\\\\Code\\\\Python\\\\Vidanges\\\\ecarts_vidanges\\\\img\\\\logo_jr.png', '.')],
    hiddenimports=[],
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
    name='excel_comparator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)