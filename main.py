#!/usr/bin/python3
#coding = utf-8

import sys
from PyQt5.QtWidgets import QApplication, QWidget
from PyQt5.QtWidgets import QApplication, QMainWindow
from calculate import *
from process import MyPyQT_Form

if __name__ == '__main__':

   app = QApplication(sys.argv)
   mainWindow = QMainWindow()
   mainWindow2 = MyPyQT_Form()
   mainWindow2.setupUi(mainWindow)
   mainWindow.show()
   sys.exit(app.exec_())
