#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :    UI.py
@Time    :    2023/05/10 22:41:34
@Author  :    cyq
@Version :    1.0
@Contact :    1135362921@qq.com
@Desc    :    
'''
import sys
import typing
import Ui_SZ
from PyQt6 import uic
from PyQt6.QtCore import Qt,QCoreApplication,QRect,QMetaObject
from qt_material import apply_stylesheet

from PyQt6.QtGui import *
from PyQt6.QtWidgets import *


class UI(QMainWindow):
    
    def __init__(self,) -> None:
        super().__init__()
               
        self.setupUi()

    def setupUi(self):
        
        menubar = self.menuBar()
        mainMenu=menubar.addMenu('主菜单(&M)')
        homePage=QAction('主页',self)
        # homePage.triggered.connect()
        SZ=QAction('水准',self)
        SZ.setShortcut('Ctrl+L')
        SZ.triggered.connect(self.LevelingSurveyingt)
        
        self.WG=QWidget(self)
        self.WG.setGeometry(50,50,300,500)
        self.Home()
        mainMenu.addActions((homePage,SZ))
        self.projectMenu=menubar.addMenu('项目管理')
        Add=QAction('新增',self)
        Add.triggered.connect(self.addP)
        self.projectMenu.addAction(Add)
        functionMenu=menubar.addMenu('功能')

        self.statusBar().showMessage('Ready')
        self.setGeometry(100, 100, 850, 650)
        self.setWindowTitle('Application')

        
        

    def Home(self):

        self.log=Ui_SZ.Ui_Form()
        self.log.setupUi(self.WG)
        QApplication.processEvents()
   
    def login(self,username,passwd):
        print(username,passwd)
        
    def LevelingSurveyingt(self):
        self.statusBar().showMessage('you clicked Button %s'%self.sender().text())
        print('hello')
                       
    def addP(self):

        self.statusBar().showMessage('you clicked Button %s'%self.sender().text())
        text, ok = QInputDialog.getText(self, 'new project','Enter your name:')
        if ok:
            self.projectMenu.addAction(text)        

def main():
    app = QApplication(sys.argv)
    apply_stylesheet(app, theme='dark_teal.xml')
    
    ex = UI()    
    ex.show()    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()