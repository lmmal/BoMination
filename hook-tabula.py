"""
PyInstaller hook for tabula-py package
"""
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Collect all data files from tabula package
datas = collect_data_files('tabula')

# Collect all submodules
hiddenimports = collect_submodules('tabula')

# Ensure Java dependencies are included
hiddenimports.extend([
    'tabula.io',
    'tabula.template',
    'tabula.util',
    'jpype1',
    'jpype1._jvmfinder'
])
