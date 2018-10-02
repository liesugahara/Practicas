# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'login.ui'
#
# Created by: PyQt5 UI code generator 5.11.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from main import *
from register import *
import pymysql

class Ui_MainWindow(QtWidgets.QWidget):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(440, 541)
        MainWindow.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(10)
        sizePolicy.setVerticalStretch(10)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        MainWindow.setSizePolicy(sizePolicy)
        MainWindow.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        MainWindow.setFocusPolicy(QtCore.Qt.NoFocus)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/Register/logo2.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        
        self.centralwidget = QtWidgets.QWidget(MainWindow)
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
        self.user_l = QtWidgets.QLineEdit(self.centralwidget)
        self.user_l.setGeometry(QtCore.QRect(19, 350, 401, 22))
        self.user_l.setObjectName("user_l")
        self.pass_l = QtWidgets.QLineEdit(self.centralwidget)
        self.pass_l.setGeometry(QtCore.QRect(20, 400, 401, 22))
        self.pass_l.setEchoMode(QtWidgets.QLineEdit.Password)
        self.pass_l.setObjectName("pass_l")
        self.register_l = QtWidgets.QPushButton(self.centralwidget)
        self.register_l.setGeometry(QtCore.QRect(80, 470, 300, 28))
        self.register_l.setObjectName("register_l")
        self.error_l = QtWidgets.QLabel(self.centralwidget)
        self.error_l.setGeometry(QtCore.QRect(80, 500, 300, 20))
        self.error_l.setText("")
        self.error_l.setObjectName("error_l")
        self.login = QtWidgets.QPushButton(self.centralwidget)
        self.login.setGeometry(QtCore.QRect(80, 430, 300, 28))
        self.login.setObjectName("login")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        
        self.login.clicked.connect(self.btn_click)#connection betweet click and slot
        self.register_l.clicked.connect(self.btn2_click)
        
#    def btn_click(self):
#        us = self.user_l.text()
#        pw = self.pass_l.text()
#        admin = 'GuHe'
#        passw = 'Cable2018'
#        if us == admin and pw == passw:
#            
#            MainWindow.close()
#            self.loginok()
#            
#        else:
##            self.error()
#            self.dbcon()
            
            
    def btn_click(self):
        us = self.user_l.text()
        pw = self.pass_l.text()
        dbcon = pymysql.connect(host="rapcwc.cakorcerkah4.us-east-2.rds.amazonaws.com", user="liesugahara", password="Cable2018", db="RAP")

        cursor = dbcon.cursor()
        cursor.execute("SELECT 'Ok' FROM RAP.login WHERE EXISTS(SELECT user, password FROM RAP.login WHERE BINARY user = BINARY %s and BINARY password = BINARY %s);", (us, pw))
        rc = cursor.rowcount
        
        if rc == 0:
            print('no ok')
            self.error()
        else:
            login_status = cursor.fetchone()
            login_status = login_status[0]
            print(login_status)

            if login_status == 'Ok':
                print('Ok')
                self.loginok()

                

        dbcon.commit()
        dbcon.close()
            
    def btn2_click(self):
        us = self.user_l.text()
        pw = self.pass_l.text()
        admin = 'GuHe'
        passw = 'Cable2018'
#        test = username
#        print(test)
        
#        a=us is admin
        if us == admin and pw == passw:
            
            self.registerok()
            MainWindow.close()            
        else:
            self.error2()
            
    def loginok(self):
#        self.b.clicked.connect(self.close)
        """ Se abre una nueva ventana"""     
#        self.l.setText('click')
        MainWindow.close()
        self.window = QtWidgets.QMainWindow()
        self.ui = Ui_MainWindow_Main()
        self.ui.setupUi_Main(self.window)
        self.window.show()
        
    def registerok(self):
#        self.b.clicked.connect(self.close)
        """ Se abre una nueva ventana"""     
#        self.l.setText('click')
        from register import Ui_MainWindow2
        self.window = QtWidgets.QMainWindow()
        self.ui = Ui_MainWindow2()
        self.ui.setup2(self.window)
        self.window.show()

    def closew(self):
        MainWindow1.close()

    def error(self):
        self.error_l.setText('Please check your username or password')
        
        
    def error2(self):
        self.error_l.setText('No tiene permisos para registrarse')

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Login"))
        self.label.setText(_translate("MainWindow", "<html><head/><body><p align=\"center\"><img src=\":/Register/logo2.png\"/></p><p align=\"center\"><img src=\":/Register/logo3.png\"/></p></body></html>"))
        self.label_2.setText(_translate("MainWindow", "<html><head/><body><p>Username</p></body></html>"))
        self.label_3.setText(_translate("MainWindow", "Password"))
        self.register_l.setText(_translate("MainWindow", "Register"))
        self.login.setText(_translate("MainWindow", "Login"))

import RAP_rc

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

