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

from PyQt6.QtWidgets import QWidget
from PyQt6.QtWidgets import QMainWindow, QApplication, QMenu
from PyQt6.QtGui import QAction
class UI(QMainWindow):
    
    def __init__(self) -> None:
        super().__init__()
        self.UImain()

    def UImain(self):

        menubar = self.menuBar()
        mainMenu=menubar.addMenu('主菜单')
        homePage=QAction('主页',self)
                
        水准=QAction('水准',self)
        mainMenu.addActions((homePage,水准))
        projectMenu=menubar.addMenu('项目管理')
        functionMenu=menubar.addMenu('功能')

        self.statusBar().showMessage('Ready')

        self.setGeometry(100, 100, 850, 650)
        self.setWindowTitle('Application')
        self.show()

        def home(self):
            
            pass

def main():
    app = QApplication(sys.argv)
    ex = UI()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()