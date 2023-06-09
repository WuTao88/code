# Form implementation generated from reading ui file 'c:\Users\Tao\Desktop\A3\code\SZ.ui'
#
# Created by: PyQt6 UI code generator 6.4.2
#
# WARNING: Any manual changes made to this file will be lost when pyuic6 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt6 import QtCore, QtGui, QtWidgets
import openpyxl as xl
import os

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(640, 480)
        self.verticalLayoutWidget = QtWidgets.QWidget(parent=Form)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(50, 40, 521, 401))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.label = QtWidgets.QLabel(parent=self.verticalLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(26)
        self.label.setFont(font)
        self.label.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label, 0, QtCore.Qt.AlignmentFlag.AlignHCenter)
        self.formLayout = QtWidgets.QFormLayout()
        self.formLayout.setFieldGrowthPolicy(QtWidgets.QFormLayout.FieldGrowthPolicy.FieldsStayAtSizeHint)
        self.formLayout.setContentsMargins(50, 15, 30, 10)
        self.formLayout.setHorizontalSpacing(20)
        self.formLayout.setVerticalSpacing(10)
        self.formLayout.setObjectName("formLayout")
        self.label_2 = QtWidgets.QLabel(parent=self.verticalLayoutWidget)
        self.label_2.setObjectName("label_2")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.ItemRole.LabelRole, self.label_2)
        self.gongcheng = QtWidgets.QComboBox(parent=self.verticalLayoutWidget)
        self.gongcheng.setObjectName("gongcheng")
        self.gongcheng.addItems(['路基工程','路面工程','绿化工程','其他'])
        self.gongcheng.setEditable(True)

        self.formLayout.setWidget(0, QtWidgets.QFormLayout.ItemRole.FieldRole, self.gongcheng)
        self.label_4 = QtWidgets.QLabel(parent=self.verticalLayoutWidget)
        self.label_4.setObjectName("label_4")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.ItemRole.LabelRole, self.label_4)
        self.path = QtWidgets.QLineEdit(parent=self.verticalLayoutWidget)
        self.path.setObjectName("path")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.ItemRole.FieldRole, self.path)
        self.label_5 = QtWidgets.QLabel(parent=self.verticalLayoutWidget)
        self.label_5.setObjectName("label_5")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.ItemRole.LabelRole, self.label_5)
        self.savePath = QtWidgets.QLineEdit(parent=self.verticalLayoutWidget)
        self.savePath.setObjectName("savePath")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.ItemRole.FieldRole, self.savePath)
        self.GD = QtWidgets.QPushButton(parent=self.verticalLayoutWidget)
        self.GD.setObjectName("GD")
        self.formLayout.setWidget(4, QtWidgets.QFormLayout.ItemRole.FieldRole, self.GD)
        self.CRD = QtWidgets.QPushButton(parent=self.verticalLayoutWidget)
        self.CRD.setObjectName("CRD")
        self.formLayout.setWidget(4, QtWidgets.QFormLayout.ItemRole.LabelRole, self.CRD)
        self.verticalLayout.addLayout(self.formLayout)
        self.tableWidget = QtWidgets.QTableWidget(parent=self.verticalLayoutWidget)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(4)
        self.tableWidget.setRowCount(12)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(7, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(8, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(9, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(10, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(11, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(3, item)
        self.verticalLayout.addWidget(self.tableWidget)

        self.retranslateUi(Form)
        self.GD.clicked.connect(lambda:Form.getData(self.tableWidget,self.path)) # type: ignore
        self.CRD.clicked.connect(lambda:Form.CreateData(self.gongcheng,self.savePath,self.tableWidget)) # type: ignore
        # QtCore.QMetaObject.connectSlotsByName(Form)

    def getData(self):
        
        filename=self.path.text()
        print(filename)
        try:
            
            wb=xl.load_workbook(f'{filename}.xlsx')
            sheet=wb.worksheets[0]
            for row in sheet.rows:            
                for cell in row:
                    self.tableWidget.setItem(cell.row-1,cell.column-1,QtWidgets.QTableWidgetItem(f'{cell.value}'))
            wb.close()
        except FileNotFoundError as err:
            print(err)
        finally:
            print('End')
            
    


    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "水准APP"))
        self.label.setText(_translate("Form", "参数输入"))
        self.label_2.setText(_translate("Form", "工程名称："))
        self.label_4.setText(_translate("Form", "参数文件路径："))
        self.path.setPlaceholderText('D:\\test.xlsx')
        self.path.setText('C:\\Users\\Tao\Desktop\\test\\tt.xlsx')
        
        self.label_5.setText(_translate("Form", "存储路径："))
        self.GD.setText(_translate("Form", "获取"))
        self.CRD.setText(_translate("Form", "生成"))
        item = self.tableWidget.verticalHeaderItem(0)
        item.setText(_translate("Form", "1"))
        item = self.tableWidget.verticalHeaderItem(1)
        item.setText(_translate("Form", "2"))
        item = self.tableWidget.verticalHeaderItem(2)
        item.setText(_translate("Form", "3"))
        item = self.tableWidget.verticalHeaderItem(3)
        item.setText(_translate("Form", "4"))
        item = self.tableWidget.verticalHeaderItem(4)
        item.setText(_translate("Form", "5"))
        item = self.tableWidget.verticalHeaderItem(5)
        item.setText(_translate("Form", "6"))
        item = self.tableWidget.verticalHeaderItem(6)
        item.setText(_translate("Form", "7"))
        item = self.tableWidget.verticalHeaderItem(7)
        item.setText(_translate("Form", "8"))
        item = self.tableWidget.verticalHeaderItem(8)
        item.setText(_translate("Form", "9"))
        item = self.tableWidget.verticalHeaderItem(9)
        item.setText(_translate("Form", "10"))
        item = self.tableWidget.verticalHeaderItem(10)
        item.setText(_translate("Form", "11"))
        item = self.tableWidget.verticalHeaderItem(11)
        item.setText(_translate("Form", "12"))
        item = self.tableWidget.horizontalHeaderItem(0)
        item.setText(_translate("Form", "桩号及部位"))
        item = self.tableWidget.horizontalHeaderItem(1)
        item.setText(_translate("Form", "桩号"))
        item = self.tableWidget.horizontalHeaderItem(2)
        item.setText(_translate("Form", "偏距"))
        item = self.tableWidget.horizontalHeaderItem(3)
        item.setText(_translate("Form", "设计高程"))
