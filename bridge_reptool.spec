# -*- mode: python ; coding: utf-8 -*-

import sys
import plistlib
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

APP_NAME = 'Bridge Report Tool'

ICON_MAC = 'icon.icns'
ICON_WIN = 'icon.ico'

# --- extras ---
datas = [
    ('bridge_defects.db', '.'),
    ('report_template.docx', '.'),
    ('inspection_report_template.docx', '.'),
]
datas += collect_data_files('docx', includes=['templates/*.xml'])

# --- Including needed modules ---
hiddenimports = (collect_submodules('docx') + ['requests'])

# --- platform-specific metadata ---
info_plist = None
bundle_identifier = None
version_file = None

if sys.platform == 'darwin':
    with open('version.plist', 'rb') as f:
        info_plist = plistlib.load(f)
        APP_NAME = info_plist.get('CFBundleName', APP_NAME)
        bundle_identifier = info_plist.get('CFBundleIdentifier')

elif sys.platform == 'win32':
    version_file = 'version.txt'

# --- Analysis ---
a = Analysis(
    ['main.py'],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

# --- Executable ---
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=APP_NAME,
    icon=ICON_MAC if sys.platform == 'darwin' else ICON_WIN,
    version=version_file,
    windowed=True,
    console=False,
)

# --- macOS bundle ---
if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name=f'{APP_NAME}.app',
        icon=ICON_MAC,
        bundle_identifier=bundle_identifier,
        info_plist=info_plist,
    )