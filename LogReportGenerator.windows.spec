# -*- mode: python ; coding: utf-8 -*-
r"""Windows build spec for LogReportGenerator.

Why a separate spec?
- Windows AV engines are more likely to flag PyInstaller --onefile builds.
- An onedir build (a dist folder) is typically faster to start and produces fewer
  false positives.

Build:
  py -m PyInstaller --noconfirm --clean LogReportGenerator.windows.spec

Output:
  dist\LogReportGenerator\LogReportGenerator.exe
"""

from PyInstaller.utils.hooks import collect_all

# Bundle source modules for dynamic imports.
datas = [('parsers', 'parsers'), ('reporting', 'reporting')]
binaries = []

hiddenimports = [
    # Parsers
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

    # Reporting
    'reporting.excel_report',
    'reporting.critical_summary',
]

# openpyxl has dynamic imports and data files.
_opx_datas, _opx_binaries, _opx_hidden = collect_all('openpyxl')
datas += _opx_datas
binaries += _opx_binaries
hiddenimports += _opx_hidden

# Hard-exclude common heavy ML stacks that PyInstaller may try to analyze if
# they are installed in the builder's global site-packages (e.g. torch).
# This tool doesn't need them; excluding keeps the build small and avoids noisy warnings.
excludes = [
  'torch',
  'torchvision',
  'tensorflow',
  'tensorboard',
  'jax',
  'jaxlib',
]


a = Analysis(
    ['gui_app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
  excludes=excludes,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

# onedir Windows executable
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='LogReportGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # UPX compression can increase AV suspicion
    console=False,
    disable_windowed_traceback=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='LogReportGenerator',
)
