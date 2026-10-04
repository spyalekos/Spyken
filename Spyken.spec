# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import copy_metadata, collect_all
from PyInstaller.building.api import Splash

block_cipher = None

datas = []
datas += copy_metadata('imageio')
datas += copy_metadata('imageio-ffmpeg')
datas += copy_metadata('moviepy')
datas += [('assets/*', 'assets')]

binaries = []
hiddenimports = [
    'imageio',
    'imageio.plugins',
    'imageio.plugins.ffmpeg',
    'imageio.v3',
    'imageio_ffmpeg',
    'numpy',
    'moviepy',
    'proglog',
    'PIL',
    'mutagen',
    'mutagen.mp3',
    'flet_desktop',
    'flet_desktop.version',
    'flet_desktop.win_taskbar',
]

for pkg in ['flet', 'flet_desktop']:
    pkg_datas, pkg_binaries, pkg_hiddenimports = collect_all(pkg)
    datas += pkg_datas
    binaries += pkg_binaries
    hiddenimports += pkg_hiddenimports

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['flet.testing', 'tkinter', 'unittest', 'pydoc'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
    optimize=0,
)

# Deduplicate binaries: remove root-level duplicate copies of DLLs that are already inside flet_desktop\app\flet\
flet_desktop_bins = {
    os.path.basename(dest.replace('/', '\\')).lower()
    for dest, src, *extra in a.binaries
    if 'flet_desktop\\app\\flet\\' in dest.replace('/', '\\')
}

a.binaries = [
    item for item in a.binaries
    if not ('\\' not in item[0].replace('/', '\\') and os.path.basename(item[0].replace('/', '\\')).lower() in flet_desktop_bins)
]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

splash = Splash(
    'assets/spyken_splash.jpg',
    binaries=a.binaries,
    datas=a.datas,
    text_pos=None,
    text_size=12,
    minify_script=True,
    always_on_top=True,
)

exe = EXE(
    pyz,
    a.scripts,
    splash,
    splash.binaries,
    a.binaries,
    a.datas,
    [],
    name='Spyken',
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
    version='version.txt',
)
