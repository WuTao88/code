#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :    Untitled-1
@Time    :    2023/05/17 15:47:11
@Author  :    cyq
@Version :    1.0
@Contact :    1135362921@qq.com
@Desc    :    
'''

import sys
import os
import SZ
import openpyxl as xl
import Ui_SZ
from qt_material import apply_stylesheet
from PyQt6.QtWidgets import *

class App(QMainWindow):

    def __init__(self) -> None:
        super().__init__()

        page=Ui_SZ.Ui_Form()
        page.setupUi(self)

    def getData(self,table,path):
        filename=path.text()
        print(filename)
        try:
            self.FileVerify(filename)
        
            wb=xl.load_workbook(f'{filename}')
            sheet=wb.worksheets[0]
            table.setRowCount(sheet.max_row)
            for row in sheet.rows:
                for cell in row:
                    table.setItem(cell.row-1,cell.column-1,QTableWidgetItem(f'{cell.value}'))
            wb.close()
        except Exception as err:
            QErrorMessage(self).showMessage(f'{err}')
           
        finally:
            
            print('End')
            
            
    def CreateData(self,gongcheng:QComboBox,savepath:QLineEdit,table:QTableWidget):

        
        print('ok')
        try:
            kwargs={'gongcheng':gongcheng.currentText(),'path':savepath.text()}
            zhs=[]
            ZHBW=''
            for row in range(table.rowCount()):
                dd=[ table.item(row,d).text() for d in range(table.columnCount())]
                
                if ZHBW==dd[0]:
                    zhs.append([ float(d) for d in dd[1:]])
                    kwargs['ZHBW']=ZHBW
                    ZHBW=dd[0]
                else:
                    if ZHBW!='':
                        print(kwargs,zhs)
                        SZ.deal(zhs,**kwargs)
                    zhs=[]
                    zhs.append([ float(d) for d in dd[1:]])
                    ZHBW=dd[0]            
            print(kwargs,zhs)
            SZ.deal(zhs,**kwargs)
        except Exception as e:
            QErrorMessage.showMessage(f'{e}')
        finally:
            QMessageBox.information(self,"信息","数据填写完成，\n欢迎下次使用。")
        pass

    def FileVerify(self,file):
        if os.path.isfile(file):
            return
        else:
            raise Exception('文件不存在')
        


if __name__=='__main__':
    app = QApplication(sys.argv)
    apply_stylesheet(app, theme='dark_purple.xml')
    ex=App()

    ex.show()


    sys.exit(app.exec())
    