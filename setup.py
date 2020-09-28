import setuptools
from cx_Freeze import setup, Executable
# Dependencies are automatically detected, but it might need
# fine tuning.

includefiles = ['template.docx','template.html']
includes = []
excludes = []
packages = []

build_options = {'packages': packages, 'excludes': excludes,'includes':includes,'include_files':includefiles}

import sys
base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('main.py', base=base, targetName = 'Automate outlook mailing')
]

setup(name='Automate outlook mailing',
      version = '1.0',
      description = 'Excel -> Word & QR -> Email',
      options = {'build_exe': build_options},
      executables = executables)

# TypeError: expected str, bytes or os.PathLike object, not NoneType
# https://stackoverflow.com/questions/62951554/cx-freeze-gives-typeerror-expected-str-bytes-or-os-pathlike-object-not-nonety
# delete the folder C:\Python38\lib\site-packages\numpy\random\_examples

# TypeError: dist must be a Distribution instance
# https://stackoverflow.com/questions/54755492/cx-freeze-typeerror-dist-must-be-a-distribution-instance
# add import setuptools before import cx_Freeze