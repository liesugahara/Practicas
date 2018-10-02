# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'register.ui'
#
# Created by: PyQt5 UI code generator 5.11.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

import pymysql


class Ui_MainWindow2(QtWidgets.QWidget):
    def setup2(self, MainWindow1):
        MainWindow1.setObjectName("MainWindow1")
        MainWindow1.setFixedSize(440, 541)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(10)
        sizePolicy.setVerticalStretch(10)
        sizePolicy.setHeightForWidth(MainWindow1.sizePolicy().hasHeightForWidth())
        MainWindow1.setSizePolicy(sizePolicy)
        MainWindow1.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        MainWindow1.setFocusPolicy(QtCore.Qt.NoFocus)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/Register/logo2.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow1.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(MainWindow1)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(60, 0, 330, 330))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(15, 330, 300, 20))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setEnabled(True)
        self.label_3.setGeometry(QtCore.QRect(15, 380, 300, 20))
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(20, 430, 300, 20))
        self.label_4.setObjectName("label_4")
        
        self.user_r = QtWidgets.QLineEdit(self.centralwidget)
        self.user_r.setGeometry(QtCore.QRect(20, 350, 401, 22))
        self.user_r.setObjectName("user_r")
        
        self.pass_r = QtWidgets.QLineEdit(self.centralwidget)
        self.pass_r.setGeometry(QtCore.QRect(20, 400, 401, 22))
        self.pass_r.setEchoMode(QtWidgets.QLineEdit.Password)
        self.pass_r.setObjectName("pass_r")
        
        
        self.cpass_r = QtWidgets.QLineEdit(self.centralwidget)
        self.cpass_r.setGeometry(QtCore.QRect(20, 450, 401, 22))
        self.cpass_r.setEchoMode(QtWidgets.QLineEdit.Password)
        self.cpass_r.setObjectName("cpass_r")
        self.register_r = QtWidgets.QPushButton(self.centralwidget)
        self.register_r.setGeometry(QtCore.QRect(70, 480, 300, 28))
        self.register_r.setObjectName("register_r")
        self.error_r = QtWidgets.QLabel(self.centralwidget)
        self.error_r.setGeometry(QtCore.QRect(70, 510, 300, 20))
        self.error_r.setText("")
        self.error_r.setObjectName("error_r")
        MainWindow1.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow1)
        QtCore.QMetaObject.connectSlotsByName(MainWindow1)
        
        
        self.register_r.clicked.connect(self.register)

    def register(self):
        username = self.user_r.text()
        passw= self.pass_r.text()
        cpassword = self.cpass_r.text()
        dbcon = pymysql.connect(host="rapcwc.cakorcerkah4.us-east-2.rds.amazonaws.com", user="liesugahara", password="Cable2018", db="RAP")
        cursor = dbcon.cursor()
        if passw == cpassword:
            cursor.execute("INSERT INTO RAP.login (`user`, `password`) VALUES (%s, %s);", (username, passw))
            dbcon.commit()
            dbcon.close()
            self.closew()
            from login import Ui_MainWindow
            self.window = QtWidgets.QMainWindow()
            self.ui = Ui_MainWindow()
            self.ui.setupUi(self.window)
            self.window.show()
            
        else:
            self.error()
            
        print(passw)
        print(cpassword)     
        
    def error(self):
        self.error_r.setText('Please check your username or password')
     
    def closew(self):
        MainWindow1.close()
        
    def retranslateUi(self, MainWindow1):
        _translate = QtCore.QCoreApplication.translate
        MainWindow1.setWindowTitle(_translate("MainWindow1", "Register"))
        self.label.setText(_translate("MainWindow1", "<html><head/><body><p align=\"center\"><img src=\":/Register/logo2.png\"/></p><p align=\"center\"><img src=\":/Register/logo3.png\"/></p></body></html>"))
        self.label_2.setText(_translate("MainWindow1", "<html><head/><body><p>Username</p></body></html>"))
        self.label_3.setText(_translate("MainWindow1", "Password"))
        self.label_4.setText(_translate("MainWindow1", "Confirm Password"))
        self.register_r.setText(_translate("MainWindow1", "Register"))

import RAP_rc

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow1 = QtWidgets.QMainWindow()
    ui = Ui_MainWindow2()
    ui.setup2(MainWindow1)
    MainWindow1.show()
    sys.exit(app.exec_())
    