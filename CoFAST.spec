# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec file for CoFAST  (onefile build → single CoFAST.exe)
# Build with:  pyinstaller CoFAST.spec

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
    ['CoFAST.py'],
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
        'test', 'tests',
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CoFAST',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,          # UPX can trigger antivirus; keep off
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,      # no terminal window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# On macOS, build a .app bundle instead
if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name='CoFAST.app',
        icon=None,              # replace with 'icon.icns' if you have one
        bundle_identifier='org.epl.cofast',
        info_plist={
            'NSHighResolutionCapable': True,
            'NSRequiresAquaSystemAppearance': False,  # allows dark mode
            'CFBundleShortVersionString': '1.0.0',
        },
    )
