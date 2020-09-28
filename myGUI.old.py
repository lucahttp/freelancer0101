import sys
from PySide2 import QtCore, QtWidgets
from PySide2.QtWidgets import QMainWindow, QWidget, QLabel, QLineEdit, QTextEdit
from PySide2.QtWidgets import QPushButton, QFileDialog
from PySide2.QtCore import QSize    

import myConfig
import myBusiness
import myExcel
class MainWindow(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)

        self.setMinimumSize(QSize(400, 250))    
        self.setWindowTitle("Automate outlook mailing")



        self.LabelExcelFile = QLabel(self)
        self.LabelExcelFile.setText('Excel :')
        self.LabelExcelFile.move(20, 20)

        self.line = QLineEdit(self)
        self.line.move(80, 20)
        self.line.resize(200, 25)


        self.searchFileButton = QPushButton("Load", self)
        self.searchFileButton.clicked.connect(self.searchAndLoadFile)
        self.searchFileButton.move(300, 20)
        self.searchFileButton.resize(75, 25)  

        # self.layout.addWidget(self.searchFileButton,2, 1, 1, 1)

        self.nameLabel = QLabel(self)
        self.nameLabel.setText('Subject :')
        self.nameLabel.move(20, 60)

        self.LineSubject = QLineEdit(self)
        self.LineSubject.move(80, 60)
        self.LineSubject.resize(200, 25)




        self.LabelArea = QLabel(self)
        self.LabelArea.setText('Body :')
        self.LabelArea.move(20, 100)

        self.text_area = QTextEdit(self)
        self.text_area.resize(200, 80)
        self.text_area.move(80, 100)


        self.pybutton = QPushButton('Save', self)
        self.pybutton.clicked.connect(self.clickMethod)
        self.pybutton.resize(200,32)
        self.pybutton.move(80, 200)


        self.createExample = QPushButton('Example File', self)
        self.createExample.clicked.connect(myExcel.CreateExample)
        self.createExample.resize(75,32)
        self.createExample.move(300, 100)



        self.autorun = QPushButton('Auto Run Off', self)
        self.autorun.clicked.connect(self.set_auto_run)
        self.autorun.resize(75,32)
        self.autorun.move(300, 150)
        self.autorun.setCheckable(True)
        
        self.start = QPushButton('Run', self)
        self.start.clicked.connect(self.run_business)
        self.start.resize(75,32)
        self.start.move(300, 200)

        
    def run_business(self):
        print("The business program starts")
        myBusiness.the_aim_of_the_program()
        pass
    
    def set_auto_run(self):
        """
        print("Select Example")
        #self.button2.click()
        self.button2.setChecked(True)
        
        if self.button4.isChecked():
            print("button4 is checked")
            self.button4.setChecked(False)
            pass
        else:
            print("button4 isnt checked")
            pass
        """
        
        if self.autorun.isChecked():
            print("autorun button is checked")
            self.autorun.setChecked(True)
            self.autorun.setText('Auto Run On')
            self.save_something('Automate outlook mailing','auto_run', 'True')
            pass
        else:
            self.autorun.setChecked(False)
            print("autorun button isnt checked")
            self.autorun.setText('Auto Run Off')
            self.save_something('Automate outlook mailing','auto_run', 'False')
            pass
        print("The business program starts")
        # myBusiness.the_aim_of_the_program()
        pass

    def save_something(self,section,somekey,something):
        if myConfig.configuration_file_has_been_persisted():
            print("The program starts")
            #print(myConfig.get_saved_data("EXCEL","subject"))
        else:
            print("The program needs to be configured")
            myConfig.configuration_file_create_persist()
            pass


        myConfig.configuration_file_set_something_to_save(section,somekey,something)
        pass
    def clickMethod(self):
        print('Your name: ' + self.line.text())


        
        if myConfig.configuration_file_has_been_persisted():
            print("The program starts")
            print(myConfig.get_saved_data("EXCEL","subject"))
        else:
            print("The program needs to be configured")
            myConfig.configuration_file_create_persist()
            pass


        if self.line.text() != "":
            myConfig.configuration_file_set_something_to_save("EXCEL","file",self.line.text())
            pass
        if self.LineSubject.text() != "":
            myConfig.configuration_file_set_something_to_save("EMAIL","subject",self.LineSubject.text())
            pass
        if self.text_area.toPlainText() != "":
            myConfig.configuration_file_set_something_to_save("EMAIL","body",self.text_area.toPlainText())
            pass
        # print(self.text_area.toPlainText())

        #print(myConfig.get_mail_data()['subject'])
        # myConfig.configuration_file_create_persist()
        pass
        """
        excel_file = "~tempfile.1.xlsx"
        #myExcel.getDataFromExcel(excel_file)

        if myConfig.configuration_file_has_been_persisted():
            print("The program starts")
            
            print(myConfig.check_data("EMAIL",'subject'))
            # for key in (myConfig.get_mail_data()): print(key)

            if myConfig.check_data("EMAIL",'subject'):
                print(myConfig.get_mail_data()['subject'])
                myConfig.get_mail_data()['subject']
                pass
            else:
                myConfig.configuration_file_create_persist()
                pass
            #myExcel.getDataFromExcel(excel_file,)
            pass
        else:
            print("The program needs to be configured")
            myConfig.configuration_file_create_persist()
            pass
        # myConfig.configuration_file_create_persist()
        pass
        """

    def setFromConfigurationFile(self):
        #excel_file = "~tempfile.1.xlsx"
        #myExcel.getDataFromExcel(excel_file)

        if myConfig.configuration_file_has_been_persisted():
            print("The program starts")
            
            #print(myConfig.check_data("EMAIL",'subject'))

            # for key in (myConfig.get_mail_data()): print(key)

            if myConfig.check_data("EMAIL",'subject'):
                #print(myConfig.get_mail_data()['subject'])
                #myConfig.get_mail_data()['subject']

                print()
                print()
                print()

                self.line.setText(myConfig.get_saved_data("EXCEL",'file'))
                # self.line.setText("Hola perro")
                self.LineSubject.setText(myConfig.get_saved_data("EMAIL",'subject'))
                self.text_area.setText(myConfig.get_saved_data("EMAIL",'body'))
                
            
                if myConfig.check_data('Automate outlook mailing','auto_run'):
                    if myConfig.get_saved_data('Automate outlook mailing','auto_run') == "True":
                        self.autorun.setChecked(True)
                        self.autorun.setText('Auto Run On')
                        pass
                    else:
                        self.autorun.setChecked(False)
                        self.autorun.setText('Auto Run Off')
                        pass
                    pass
                else:
                    print('auto_run not exist')
                    myConfig.configuration_file_create_persist()
                    pass
                pass
            else:
                myConfig.configuration_file_create_persist()
                pass
            #myExcel.getDataFromExcel(excel_file,)
            pass
        else:
            print("The program needs to be configured")
            myConfig.configuration_file_create_persist()
            pass

        pass


    def searchAndLoadFile(self):
        #path_to_file, _ = QFileDialog.getOpenFileName(self, self.tr("Load Image"), self.tr("~/Desktop/"), self.tr("Images (*.jpg)"))
        path_to_file, _  = QFileDialog.getOpenFileName(self, self.tr("Load Excel"), self.tr("~/Desktop/"), self.tr("/ (*.xlsx)"))

        #self.test(path_to_file)
        print(path_to_file)

        #self.testFuncion(path_to_file)

        # self.filenameLoaded = path_to_file
        
        self.line.setText(path_to_file)

"""
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    mainWin = MainWindow()
    mainWin.setFromConfigurationFile()
    mainWin.show()
    sys.exit( app.exec_() )
"""