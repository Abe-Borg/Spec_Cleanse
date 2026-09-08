# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller build for SpecCleanse.

Produces a single windowed executable. Build it from the project root:

    pyinstaller packaging/speccleanse.spec --noconfirm

PyInstaller does not cross-compile — a Windows .exe has to be built on Windows.
.github/workflows/release.yml does that on a windows runner.
"""

import os

# SPECPATH is the directory holding this file; the sources sit one level up.
PROJECT_ROOT = os.path.abspath(os.path.join(SPECPATH, os.pardir))


a = Analysis(
    [os.path.join(PROJECT_ROOT, "gui.py")],
    pathex=[PROJECT_ROOT],
    binaries=[],
    # The default patterns.yaml ships inside the bundle. It is not what the user
    # edits: apppaths.resolve_config_path copies it out to %APPDATA% on first run,
    # because PyInstaller deletes this extraction directory when the process ends.
    datas=[(os.path.join(PROJECT_ROOT, "patterns.yaml"), ".")],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Nothing here is imported by SpecCleanse; excluding them keeps the binary
    # from picking up whatever else happens to be in the build environment.
    excludes=[
        "numpy",
        "matplotlib",
        "PIL",
        "pytest",
        "setuptools",
        "pip",
    ],
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
    name="SpecCleanse",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # UPX compression is left off deliberately. It saves a few MB and is a
    # reliable way to get an unsigned .exe flagged by corporate antivirus,
    # which matters more than the size here.
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    # Tkinter app: no console window should appear behind it.
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
