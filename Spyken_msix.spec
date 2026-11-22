# -*- mode: python ; coding: utf-8 -*-
"""Slimmed PyInstaller onedir spec for Spyken MSIX (Microsoft Store)."""
import os
from PyInstaller.utils.hooks import collect_data_files, copy_metadata

datas = []

# Runtime visual & avatar assets (MSIX uses its own Assets\SplashScreen.png,
# while spyken_splash.jpg / icon.ico / avatars are kept for app UI & video rendering)
for fname in (
    "icon.png",
    "icon.ico",
    "teacher_idle.png",
    "teacher_talk.png",
    "teacher_talk2.png",
    "female_teacher_idle.png",
    "female_teacher_talk.png",
    "female_teacher_talk2.png",
):
    fpath = os.path.join("assets", fname)
    if os.path.exists(fpath):
        datas.append((fpath, "assets"))

# Targeted Flet runtime data only (icons JSON — NO .pyi stubs or CLI bloat)
datas += collect_data_files("flet", includes=["**/*.json"])

# Metadata & bundled FFmpeg binary for audio post-processing and MoviePy rendering
datas += copy_metadata("imageio")
datas += copy_metadata("imageio-ffmpeg")
datas += copy_metadata("moviepy")
datas += collect_data_files("imageio_ffmpeg", includes=["binaries/*"])
datas += collect_data_files("certifi")

hiddenimports = [
    "flet",
    "flet_desktop",
    "flet_desktop.extensions",
    "flet_desktop.version",
    "flet_desktop.win_taskbar",
    "edge_tts",
    "edge_tts.drm",
    "aiohttp",
    "certifi",
    "docx",
    "lxml",
    "lxml.etree",
    "lxml._elementpath",
    "pymupdf",
    "fitz",
    "moviepy",
    "imageio",
    "imageio.plugins",
    "imageio.plugins.ffmpeg",
    "imageio.v3",
    "imageio_ffmpeg",
    "numpy",
    "proglog",
    "mutagen",
    "mutagen.mp3",
    "PIL.Image",
    "PIL.ImageDraw",
    "PIL.ImageFont",
    "PIL.PngImagePlugin",
    "PIL.JpegImagePlugin",
    "PIL.IcoImagePlugin",
    "PIL.BmpImagePlugin",
]

excludes = [
    # Dev / Build / Packaging / Test / REPL tools
    "setuptools",
    "pkg_resources",
    "wheel",
    "_distutils_hack",
    "pip",
    "_pyrepl",
    "unittest",
    "doctest",
    "pdb",
    "pydoc",
    "test",
    "watchdog",
    # Unused GUI frameworks
    "tkinter",
    "_tkinter",
    # Heavy transitive CLI/auth deps not used at runtime
    "pygments",
    "markdown_it",
    "mdurl",
    "oauthlib",
    # Unused Flet CLI / Test / Auth / Canvas submodules
    "flet.cli",
    "flet.testing",
    "flet.pytest_plugin",
    "flet.auth",
    "flet.canvas",
    "flet.utils.griffe_deprecations",
    # Unused heavy Pillow plugins (KEEP PIL.ImageFont / _imagingft for DOCX & Marquee video text!)
    "PIL.AvifImagePlugin",
    "PIL.WebPImagePlugin",
    "PIL.ImageCms",
    "PIL.SpiderImagePlugin",
    # Unused heavy optional packages pulled by pymupdf 2.x or transitive deps
    "onnxruntime",
    "pymupdf_layout",
    "pymupdf4llm",
    "matplotlib",
    "scipy",
    "pandas",
    "IPython",
    "jupyter",
]

a = Analysis(
    ["main.py"],
    pathex=["."],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

# Remove:
# 1. Duplicate Flet client files injected by flet_cli's hook-flet.py (MSIX bundles ONE clean copy at flet_client/)
# 2. Unused python-docx dev template folder (contains OPC-illegal `[Content_Types].xml` that breaks makeappx.exe 0x8007007b)
# 3. Unused pymupdf ONNX layout models & C++ devel headers/libs (~55 MB of dead weight)
_DROP_DATA_PREFIXES = (
    "flet_desktop/app",
    "docx/templates/default-docx-template",
    "pymupdf/layout",
    "pymupdf/mupdf-devel",
)
a.datas = [
    item for item in a.datas
    if not any(item[0].replace("\\", "/").lower().startswith(pref) for pref in _DROP_DATA_PREFIXES)
]

# Drop unused heavy Pillow C-extensions (KEEP _imagingft!) AND duplicate Flutter DLLs leaked from hook-flet.py
_DROP_BIN_SUBSTRINGS = (
    "_avif",
    "_webp",
    "_imagingcms",
    "flet_desktop/app",
    "flet_desktop\\app",
    "libmpv-2",
    "flutter_windows",
    "libegl.dll",
    "libglesv2.dll",
    "d3dcompiler_47",
    "vk_swiftshader",
    "vulkan-1",
    "rive_native",
    "_plugin.dll",
    "dartjni",
    "flutter_local_notifications",
    "flet.exe",
)
a.binaries = [
    item for item in a.binaries
    if not any(sub in item[0].lower() for sub in _DROP_BIN_SUBSTRINGS)
]

pyz = PYZ(a.pure)

icon_file = "assets/icon.ico" if os.path.exists("assets/icon.ico") else None

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Spyken",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_file,
    version="version.txt",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="Spyken_msix",
)
