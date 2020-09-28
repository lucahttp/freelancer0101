

# https://stackoverflow.com/questions/14165398/a-good-python-to-exe-compiler
# https://cx-freeze.readthedocs.io/en/latest/installation.html
```
pip install cx_Freeze --upgrade
```

```
cxfreeze-quickstart
```


´´´
import setuptools
from cx_Freeze import setup, Executable
# Dependencies are automatically detected, but it might need
# fine tuning.

includefiles = ['README.txt', 'template_2.docx']
includes = []
excludes = []
packages = []

build_options = {'packages': [], 'excludes': [],'includes':includes,'include_files':includefiles}

import sys
if sys.platform == 'win32':
    base = 'Win32GUI'
base = 'Console'

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

´´´


UPX




# for %v in (*.*) do C:\Users\Luca\Desktop\upx-3.96-win64\upx.exe "C:\Users\Luca\Desktop\Automate outlook mailing\%v"



# https://stackoverflow.com/questions/44544369/i-am-not-able-to-add-an-image-in-email-body-using-python-i-am-able-to-add-a-pi/44619761
"""
# attachment = mail.Attachments.Add("C:\Users\MA299445\Downloads\screenshot.png")
# attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
# mail.HTMLBody = "<html><body>Test image <img src=""cid:MyId1""></body></html>"
"""


to compile:
run powershell as admin
cd to/the/path/of/the/project
C:/Python38/python.exe  setup.py install