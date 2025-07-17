# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_submodules

datas = [('C:\\Users\\Luke Malkasian\\Documents\\OMNI\\BoMination\\src\\chromedriver.exe', '.'), ('C:\\Users\\Luke Malkasian\\Documents\\OMNI\\BoMination\\Files\\OCTF-1539-COST SHEET.xlsx', 'Files'), ('C:\\Users\\Luke Malkasian\\Documents\\OMNI\\BoMination\\src', 'src'), ('C:\\Users\\Luke Malkasian\\Documents\\OMNI\\BoMination\\assets', 'assets')]
hiddenimports = ['ttkbootstrap', 'selenium', 'pandas', 'numpy', 'matplotlib', 'openpyxl', 'xlrd', 'tabula', 'pandastable', 'PIL', 'tkinter', 'jpype1', 'jpype1._jvmfinder', 'packaging', 'packaging.version', 'packaging.specifiers', 'packaging.requirements', 'pipeline', 'pipeline.main_pipeline', 'pipeline.extract_main', 'pipeline.extract_bom_tab', 'pipeline.extract_bom_cam', 'pipeline.lookup_price', 'pipeline.map_cost_sheet', 'pipeline.validation_utils', 'pipeline.ocr_preprocessor', 'gui', 'gui.review_window', 'gui.roi_picker', 'gui.table_selector', 'omni_cust', 'omni_cust.customer_config', 'omni_cust.customer_formatters']
datas += collect_data_files('ttkbootstrap')
datas += collect_data_files('tabula')
hiddenimports += collect_submodules('selenium')
hiddenimports += collect_submodules('ttkbootstrap')
hiddenimports += collect_submodules('pandas')
hiddenimports += collect_submodules('tabula')
hiddenimports += collect_submodules('packaging')
hiddenimports += collect_submodules('pipeline')
hiddenimports += collect_submodules('gui')
hiddenimports += collect_submodules('omni_cust')


a = Analysis(
    ['C:\\Users\\Luke Malkasian\\Documents\\OMNI\\BoMination\\BoMinationApp_wrapper.py'],
    pathex=['C:\\Users\\Luke Malkasian\\Documents\\OMNI\\BoMination\\src', 'C:\\Users\\Luke Malkasian\\Documents\\OMNI\\BoMination\\src/pipeline', 'C:\\Users\\Luke Malkasian\\Documents\\OMNI\\BoMination\\src/gui', 'C:\\Users\\Luke Malkasian\\Documents\\OMNI\\BoMination\\src/omni_cust'],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=['.'],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib.tests', 'numpy.tests', 'pandas.tests', 'PIL.tests'],
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
    name='BoMinationApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['C:\\Users\\Luke Malkasian\\Documents\\OMNI\\BoMination\\assets\\BoMination_black.ico'],
)
