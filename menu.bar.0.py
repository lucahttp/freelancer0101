import sys
from PySide2.QtCore import *
from PySide2.QtGui import *
from PySide2 import QtCore, QtWidgets
from PySide2.QtWidgets import QMainWindow, QWidget, QLabel, QLineEdit, QTextEdit
from PySide2.QtWidgets import QPushButton, QFileDialog
from PySide2.QtCore import QSize    

class menudemo(QMainWindow):
   def __init__(self, parent = None):
      super(menudemo, self).__init__(parent)
		
      layout = QtWidgets.QHBoxLayout()
      bar = self.menuBar()
      file = bar.addMenu("File")
      file.addAction("New")
		
      save = QtWidgets.QAction("Save",self)
      save.setShortcut("Ctrl+S")
      file.addAction(save)
		
      edit = file.addMenu("Edit")
      edit.addAction("copy")
      edit.addAction("paste")
		
      quit = QtWidgets.QAction("Quit",self) 
      file.addAction(quit)
      file.triggered[QtWidgets.QAction].connect(self.processtrigger)
      self.setLayout(layout)
      self.setWindowTitle("menu demo")
		
   def processtrigger(self,q):
      print(q.text()+" is triggered")
		
def main():
   app = QtWidgets.QApplication(sys.argv)
   ex = menudemo()
   ex.show()
   sys.exit(app.exec_())
	
if __name__ == '__main__':
   main()