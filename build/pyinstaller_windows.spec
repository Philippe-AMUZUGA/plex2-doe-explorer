# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules

root = Path.cwd()
hiddenimports = collect_submodules("pyDOE3")
app_dir = root / "app"
assets_dir = app_dir / "assets"

block_cipher = None

a = Analysis(
    [str(app_dir / "PLEX2_Launcher.py")],
    pathex=[str(app_dir)],
    binaries=[],
    datas=[
        (str(assets_dir / "plex2_icon.ico"), "assets"),
        (str(assets_dir / "plex2_icon.png"), "assets"),
    ],
    hiddenimports=hiddenimports,
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
    name='PLEX2',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=str(assets_dir / "plex2_icon.ico"),
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PLEX2',
)
