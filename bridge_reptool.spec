# -*- mode: python ; coding: utf-8 -*-

import sys
import plistlib
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

APP_NAME = "BridgeReportTool"

ICON_MAC = "icon.icns"
ICON_WIN = "icon.ico"

datas = [
    ("bridge_defects.db", "."),
    ("report_template.docx", "."),
    ("inspection_report_template.docx", "."),
]

# забираем docx/templates/*.xml и прочие data-файлы python-docx
datas += collect_data_files("docx")

hiddenimports = collect_submodules("docx") + ["requests"]

info_plist = None
bundle_identifier = None
version_file = None

if sys.platform == "darwin":
    with open("version.plist", "rb") as f:
        info_plist = plistlib.load(f)
        APP_NAME = info_plist.get("CFBundleName", APP_NAME)
        bundle_identifier = info_plist.get("CFBundleIdentifier")
elif sys.platform == "win32":
    version_file = "version.txt"

a = Analysis(
    ["main.py"],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=True, 
)

pyz = PYZ(a.pure)

# onedir: в EXE НЕ кладём zipfiles/datas, они идут в COLLECT
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=APP_NAME,
    icon=ICON_MAC if sys.platform == "darwin" else ICON_WIN,
    version=version_file,
    windowed=True,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    name=APP_NAME,
)

if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name=f"{APP_NAME}.app",
        icon=ICON_MAC,
        bundle_identifier=bundle_identifier,
        info_plist=info_plist,
    )