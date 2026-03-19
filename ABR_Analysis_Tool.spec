# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec file for ABR Analysis Tool
# Build with:  pyinstaller ABR_Analysis_Tool.spec

import sys
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

# Collect all scipy and matplotlib submodules to avoid missing-module errors
hidden_imports = (
    collect_submodules('scipy')
    + collect_submodules('scipy.optimize')
    + collect_submodules('scipy.signal')
    + collect_submodules('matplotlib')
    + collect_submodules('openpyxl')
    + [
        'matplotlib.backends.backend_tkagg',
        'matplotlib.backends._backend_tk',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'pandas',
        'numpy',
    ]
)

datas = (
    collect_data_files('matplotlib')
    + collect_data_files('scipy')
)

a = Analysis(
    ['abr_analysis_tool.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'PyQt5', 'PyQt6', 'PySide2', 'PySide6',
        'wx', 'gi', 'gtk',
        'IPython', 'jupyter', 'notebook',
        'test', 'tests', 'unittest',
    ],
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
    name='ABR Analysis Tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,          # UPX can trigger antivirus; keep off
    console=False,      # no terminal window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='ABR Analysis Tool',
)

# On macOS, also build a .app bundle
if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='ABR Analysis Tool.app',
        icon=None,              # replace with 'icon.icns' if you have one
        bundle_identifier='org.epl.abr-analysis-tool',
        info_plist={
            'NSHighResolutionCapable': True,
            'NSRequiresAquaSystemAppearance': False,  # allows dark mode
            'CFBundleShortVersionString': '1.0.0',
        },
    )
