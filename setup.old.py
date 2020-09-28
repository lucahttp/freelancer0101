import setuptools
from cx_Freeze import setup, Executable
# Dependencies are automatically detected, but it might need
# fine tuning.

includefiles = ['README.txt', 'template_2.docx']
includes = []
excludes = []
packages = []

build_options = {'packages': packages, 'excludes': excludes,'includes':includes,'include_files':includefiles}

import sys
if sys.platform == 'win32':
    base = 'Win32GUI'
base = 'Console'
base = 'Win32GUI'

executables = [
    Executable('main.py', base=base)
]

setup(name="Automate outlook mailing",
      version = '1.0',
      description = '',
      options = {'build_exe': build_options},
      executables = executables)

# TypeError: expected str, bytes or os.PathLike object, not NoneType
# https://stackoverflow.com/questions/62951554/cx-freeze-gives-typeerror-expected-str-bytes-or-os-pathlike-object-not-nonety
# delete the folder C:\Python38\lib\site-packages\numpy\random\_examples

# TypeError: dist must be a Distribution instance
# https://stackoverflow.com/questions/54755492/cx-freeze-typeerror-dist-must-be-a-distribution-instance
# add import setuptools before import cx_Freeze