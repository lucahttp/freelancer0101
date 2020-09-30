# freelancer0101
 
# C:/Python38/python.exe -m pip install auto-py-to-exe
# C:/Python38/python.exe -m auto-py-to-exe
# py -3.7 -m pyinstaller --noconfirm --onedir --console --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/main.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myBusiness.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myConfig.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myExcel.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myGUI.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myMail.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myQRCode.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myWord.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/template_2.docx;."  "C:/Users/Luca/Documents/GitHub/freelancer01/main.py"
# pyinstaller --noconfirm --onedir --console --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/main.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myBusiness.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myConfig.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myExcel.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myGUI.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myMail.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myQRCode.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/myWord.py;." --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/template_2.docx;."  "C:/Users/Luca/Documents/GitHub/freelancer01/main.py"

# works
# pyinstaller --noconfirm --onedir --console --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/template_2.docx;."  "C:/Users/Luca/Documents/GitHub/freelancer01/main.py"

# pyinstaller --noconfirm --onefile --console --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/template_2.docx;."  "C:/Users/Luca/Documents/GitHub/freelancer01/main.py"

# pip freeze > requirements.txt
# ModuleNotFoundError: No module named 'Image'
#  pip install pillow


# pyinstaller --noconfirm --onedir --console --add-data "C:/Users/Luca/Documents/GitHub/freelancer01/template_2.docx;."  "C:/Users/Luca/Documents/GitHub/freelancer01/main.py"
"""


pyinstaller --noconfirm --onedir --console --upx-dir "C:/upx-3.96-win64" --add-data "E:/GitHub/freelancer0101/template.docx;." --add-data "E:/GitHub/freelancer0101/template.html;."  "E:/GitHub/freelancer0101/main.py"

'Automate outlook mailing'
openpyxl

PySide2


docx docxtpl mailmerge
qrcode pillow
"""


pyinstaller --noconfirm --onedir --console --name "Automate outlook mailing" --upx-dir "C:/upx-3.96-win64" --add-data "E:/GitHub/freelancer0101/template.docx;." --add-data "E:/GitHub/freelancer0101/template.html;."  "E:/GitHub/freelancer0101/main.py"



ERROR:

pywintypes.com_error: (-2147221008, 'No se ha llamado a CoInitialize.', None, None)

solution:

https://yiruiscool.wordpress.com/2015/04/29/activate-initialize-com-library-for-calling-thread-in-python/