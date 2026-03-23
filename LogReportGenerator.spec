# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('parsers', 'parsers'), ('reporting', 'reporting')]
binaries = []
# Keep this list explicit: it avoids packaged builds missing modules that are
# imported dynamically via the registry or plugin system.
hiddenimports = [
    'parsers.base',
    'parsers.registry',
    'parsers.ansys',
    'parsers.ansys_peak',
    'parsers.catia_license',
    'parsers.catia_token',
    'parsers.catia_usage_stats',
    'parsers.cortona',
    'parsers.cortona_admin',
    'parsers.creo',
    'parsers.matlab',
    'parsers.nx',
    'reporting.excel_report',
    'reporting.critical_summary',
]
tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['gui_app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='LogReportGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='LogReportGenerator',
)
app = BUNDLE(
    coll,
    name='LogReportGenerator.app',
    icon=None,
    bundle_identifier=None,
)
