# Form implementation generated from reading ui file 'c:\Users\Tao\Desktop\A3\code\Login.ui'
#
# Created by: PyQt6 UI code generator 6.4.2
#
# WARNING: Any manual changes made to this file will be lost when pyuic6 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt6 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(640, 480)
        self.verticalLayoutWidget = QtWidgets.QWidget(parent=Form)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(70, 50, 471, 371))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.label = QtWidgets.QLabel(parent=self.verticalLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("华文行楷")
        font.setPointSize(36)
        self.label.setFont(font)
        self.label.setWordWrap(False)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label, 0, QtCore.Qt.AlignmentFlag.AlignHCenter)
        self.formLayout = QtWidgets.QFormLayout()
        self.formLayout.setFieldGrowthPolicy(QtWidgets.QFormLayout.FieldGrowthPolicy.FieldsStayAtSizeHint)
        self.formLayout.setRowWrapPolicy(QtWidgets.QFormLayout.RowWrapPolicy.DontWrapRows)
        self.formLayout.setContentsMargins(20, 15, 10, -1)
        self.formLayout.setSpacing(20)
        self.formLayout.setObjectName("formLayout")
        self.label_2 = QtWidgets.QLabel(parent=self.verticalLayoutWidget)
        self.label_2.setObjectName("label_2")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.ItemRole.LabelRole, self.label_2)
        self.user = QtWidgets.QLineEdit(parent=self.verticalLayoutWidget)
        self.user.setObjectName("lineEdit")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.ItemRole.FieldRole, self.user)
        self.label_3 = QtWidgets.QLabel(parent=self.verticalLayoutWidget)
        self.label_3.setObjectName("label_3")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.ItemRole.LabelRole, self.label_3)
        self.passwd = QtWidgets.QLineEdit(parent=self.verticalLayoutWidget)
        self.passwd.setObjectName("lineEdit_2")
        self.passwd.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.ItemRole.FieldRole, self.passwd)
        self.pushButton = QtWidgets.QPushButton(parent=self.verticalLayoutWidget)
        font = QtGui.QFont()
        font.setFamily("华文仿宋")
        font.setPointSize(36)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.ItemRole.FieldRole, self.pushButton)
        self.verticalLayout.addLayout(self.formLayout)

        self.retranslateUi(Form)
        
        self.pushButton.clicked.connect(lambda:Form.login(self.user.text(),self.passwd.text())) # type: ignore
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label.setText(_translate("Form", "欢迎使用水准小程序"))
        self.label_2.setText(_translate("Form", "用户名："))
        self.user.setPlaceholderText(_translate("Form", "请输入用户名"))
        self.label_3.setText(_translate("Form", "密  码："))
        self.passwd.setPlaceholderText(_translate("Form", "请输入密码"))
        self.pushButton.setText(_translate("Form", "登录"))
