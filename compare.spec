# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['compare.py'],
    pathex=['C:/Users/Administrator/AppData/Local/Programs/Python/Python38/Lib/site-packages'],
    binaries=[],
    datas=[],
    hiddenimports=['pyexcel_xls', 'pyexcel_xlsx', 'pyexcel_xlsxw', 'pyexcel_io.readers.csv_in_file', 'pyexcel_io.readers.csv_in_memory', 'pyexcel_io.readers.csv_content', 'pyexcel_io.readers.csvz', 'pyexcel_io.writers.csv_in_file', 'pyexcel_io.writers.csv_in_memory', 'pyexcel_io.writers.csvz_writer', 'pyexcel_io.database.importers.django', 'pyexcel_io.database.importers.sqlalchemy', 'pyexcel_io.database.exporters.django', 'pyexcel_io.database.exporters.sqlalchemy'],
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
    a.binaries,
    a.datas,
    [],
    name='compare',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
