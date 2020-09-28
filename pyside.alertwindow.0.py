import sys
from PySide2 import QtGui, QtCore
from PySide2.QtCore import *
from PySide2.QtGui import *
from PySide2 import QtCore, QtWidgets
from PySide2.QtWidgets import QMainWindow, QWidget, QLabel, QLineEdit, QTextEdit
from PySide2.QtWidgets import QPushButton, QFileDialog
from PySide2.QtCore import QSize    
class Window(QMainWindow):

    def __init__(self):
        super(Window, self).__init__()
        self.setGeometry(50, 50, 500, 300)
        self.setWindowTitle("PyQT tuts!")
        self.setWindowIcon(QtGui.QIcon('pythonlogo.png'))

        extractAction = QtWidgets.QAction("&GET TO THE CHOPPAH!!!", self)
        extractAction.setShortcut("Ctrl+Q")
        extractAction.setStatusTip('Leave The App')
        extractAction.triggered.connect(self.close_application)

        self.statusBar()

        mainMenu = self.menuBar()
        fileMenu = mainMenu.addMenu('&File')
        fileMenu.addAction(extractAction)
        

        self.home()

    def home(self):
        btn = QtWidgets.QPushButton("Quit", self)
        btn.clicked.connect(self.close_application)
        btn.resize(btn.minimumSizeHint())
        btn.move(0,100)

        extractAction = QtWidgets.QAction(QtGui.QIcon('todachoppa.png'), 'Flee the Scene', self)
        extractAction.triggered.connect(self.close_application)
        
        self.toolBar = self.addToolBar("Extraction")
        self.toolBar.addAction(extractAction)
 
        self.show()

    def close_application(self):
        choice = QtWidgets.QMessageBox.question(self, 'Extract!',
                                            "Get into the chopper?",
                                            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if choice == QtWidgets.QMessageBox.Yes:
            print("Extracting Naaaaaaoooww!!!!")
            sys.exit()
        else:
            pass

    
def run():
    app = QtWidgets.QApplication(sys.argv)
    GUI = Window()
    sys.exit(app.exec_())


run()