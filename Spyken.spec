# -*- mode: python ; coding: utf-8 -*-
"""Slimmed PyInstaller onefile spec for standalone Spyken.exe."""
import os
import sys
from PyInstaller.utils.hooks import collect_data_files, copy_metadata

if sys.platform == "win32":
    from PyInstaller.building.api import Splash

block_cipher = None

datas = []

# Only include runtime visual & avatar assets (exclude ms-icon.png & duplicate spyken_splash.jpg)
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

# Targeted Flet runtime JSON only (icons.json, cupertino_icons.json — NO .pyi stubs or CLI bloat)
datas += collect_data_files("flet", includes=["**/*.json"])

# Metadata & bundled FFmpeg / SSL certs for audio/video rendering and Edge-TTS
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
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
    optimize=0,
)

# Drop unused python-docx dev template & pymupdf ONNX layout models / devel headers
# IMPORTANT: Keep `flet_desktop/app/` for standalone onefile execution without runtime downloads!
_DROP_DATA_PREFIXES = (
    "docx/templates/default-docx-template",
    "pymupdf/layout",
    "pymupdf/mupdf-devel",
)
a.datas = [
    item for item in a.datas
    if not any(item[0].replace("\\", "/").lower().startswith(pref) for pref in _DROP_DATA_PREFIXES)
]

# Drop unused heavy Pillow C-extensions (KEEP _imagingft!) AND root-level duplicate Flutter binaries
_DROP_PIL_SUBSTRINGS = ("_avif", "_webp", "_imagingcms")
_DUP_FLUTTER_SUBSTRINGS = (
    "libmpv",
    "flutter_windows",
    "libflutter_linux",
    "libegl",
    "libglesv2",
    "d3dcompiler_47",
    "vk_swiftshader",
    "vulkan-1",
    "rive_native",
    "_plugin.dll",
    "_plugin.so",
    "dartjni",
    "flutter_local_notifications",
    "flet.exe",
)


def _keep_binary(entry):
    dest_norm = entry[0].replace("\\", "/").lower()
    if any(sub in dest_norm for sub in _DROP_PIL_SUBSTRINGS):
        return False
    # Keep the single copy under flet_desktop/app/ and strip root-level duplicates
    if not dest_norm.startswith("flet_desktop/app/") and any(
        sub in dest_norm for sub in _DUP_FLUTTER_SUBSTRINGS
    ):
        return False
    return True


a.binaries = [item for item in a.binaries if _keep_binary(item)]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

icon_file = "assets/icon.ico" if os.path.exists("assets/icon.ico") else None

splash_args = []
if sys.platform == "win32":
    splash = Splash(
        "assets/spyken_splash.jpg",
        binaries=a.binaries,
        datas=a.datas,
        text_pos=None,
        text_size=12,
        minify_script=True,
        always_on_top=True,
    )
    splash_args = [splash, splash.binaries]

exe = EXE(
    pyz,
    a.scripts,
    *splash_args,
    a.binaries,
    a.datas,
    [],
    name="Spyken",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_file,
    version="version.txt",
)
