# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'main2.ui'
#
# Created by: PyQt5 UI code generator 5.11.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
import xlrd
import re
import numpy as np
import datetime
import xlwt

class Ui_MainWindow(QtWidgets.QWidget):
    def setupUi(self, MainWindow):
        """
# =============================================================================
#         DO NOT MODIFY - GUI CORE
# =============================================================================
        """
        MainWindow.setObjectName("MainWindow")
        MainWindow.setEnabled(True)
        MainWindow.resize(1580, 1020)
        MainWindow.setDockNestingEnabled(False)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.scrollArea = QtWidgets.QScrollArea(self.centralwidget)
        self.scrollArea.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setObjectName("scrollArea")
        self.scrollAreaWidgetContents = QtWidgets.QWidget()
        self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 1558, 947))
        self.scrollAreaWidgetContents.setObjectName("scrollAreaWidgetContents")
        self.gridLayout = QtWidgets.QGridLayout(self.scrollAreaWidgetContents)
        self.gridLayout.setObjectName("gridLayout")
        self.mrc_sf = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.mrc_sf.setMinimumSize(QtCore.QSize(221, 31))
        self.mrc_sf.setMaximumSize(QtCore.QSize(221, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.mrc_sf.setFont(font)
        self.mrc_sf.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.mrc_sf.setText("")
        self.mrc_sf.setAlignment(QtCore.Qt.AlignCenter)
        self.mrc_sf.setObjectName("mrc_sf")
        self.gridLayout.addWidget(self.mrc_sf, 23, 2, 1, 1)
        self.label_3 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_3.setMinimumSize(QtCore.QSize(151, 16))
        self.label_3.setMaximumSize(QtCore.QSize(151, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.gridLayout.addWidget(self.label_3, 18, 0, 1, 1)
        self.id_sf = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.id_sf.setMinimumSize(QtCore.QSize(221, 31))
        self.id_sf.setMaximumSize(QtCore.QSize(221, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.id_sf.setFont(font)
        self.id_sf.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.id_sf.setText("")
        self.id_sf.setAlignment(QtCore.Qt.AlignCenter)
        self.id_sf.setObjectName("id_sf")
        self.gridLayout.addWidget(self.id_sf, 20, 2, 1, 1)
        self.exportar_sf = QtWidgets.QPushButton(self.scrollAreaWidgetContents)
        self.exportar_sf.setMinimumSize(QtCore.QSize(93, 28))
        self.exportar_sf.setMaximumSize(QtCore.QSize(93, 28))
        self.exportar_sf.setObjectName("exportar_sf")
        self.gridLayout.addWidget(self.exportar_sf, 19, 3, 1, 1)
        self.label_29 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_29.setMinimumSize(QtCore.QSize(71, 31))
        self.label_29.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_29.setFont(font)
        self.label_29.setAlignment(QtCore.Qt.AlignCenter)
        self.label_29.setObjectName("label_29")
        self.gridLayout.addWidget(self.label_29, 20, 1, 1, 1)
        self.select_sf = QtWidgets.QComboBox(self.scrollAreaWidgetContents)
        self.select_sf.setMinimumSize(QtCore.QSize(151, 22))
        self.select_sf.setMaximumSize(QtCore.QSize(151, 22))
        self.select_sf.setObjectName("select_sf")
        self.select_sf.addItem("")
        self.select_sf.setItemText(0, "")
        self.select_sf.addItem("")
        self.select_sf.addItem("")
        self.select_sf.addItem("")
        self.select_sf.addItem("")
        self.gridLayout.addWidget(self.select_sf, 20, 3, 1, 1)
        self.label_35 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_35.setMinimumSize(QtCore.QSize(71, 31))
        self.label_35.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_35.setFont(font)
        self.label_35.setStyleSheet("")
        self.label_35.setAlignment(QtCore.Qt.AlignCenter)
        self.label_35.setObjectName("label_35")
        self.gridLayout.addWidget(self.label_35, 19, 1, 1, 1)
        self.label_33 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_33.setMinimumSize(QtCore.QSize(71, 31))
        self.label_33.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_33.setFont(font)
        self.label_33.setAlignment(QtCore.Qt.AlignCenter)
        self.label_33.setObjectName("label_33")
        self.gridLayout.addWidget(self.label_33, 21, 1, 1, 1)
        self.operador_sf = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.operador_sf.setMinimumSize(QtCore.QSize(221, 40))
        self.operador_sf.setMaximumSize(QtCore.QSize(221, 40))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.operador_sf.setFont(font)
        self.operador_sf.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.operador_sf.setText("")
        self.operador_sf.setAlignment(QtCore.Qt.AlignCenter)
        self.operador_sf.setObjectName("operador_sf")
        self.gridLayout.addWidget(self.operador_sf, 21, 2, 1, 1)
        self.tabla_sf = QtWidgets.QTableWidget(self.scrollAreaWidgetContents)
        self.tabla_sf.setMinimumSize(QtCore.QSize(1031, 280))
        self.tabla_sf.setMaximumSize(QtCore.QSize(16777215, 280))
        self.tabla_sf.setEditTriggers(QtWidgets.QAbstractItemView.AnyKeyPressed|QtWidgets.QAbstractItemView.DoubleClicked|QtWidgets.QAbstractItemView.EditKeyPressed|QtWidgets.QAbstractItemView.SelectedClicked)
        self.tabla_sf.setObjectName("tabla_sf")
        self.tabla_sf.setColumnCount(0)
        self.tabla_sf.setRowCount(0)
        self.gridLayout.addWidget(self.tabla_sf, 19, 0, 8, 1)
        self.norden_sf = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.norden_sf.setMinimumSize(QtCore.QSize(221, 31))
        self.norden_sf.setMaximumSize(QtCore.QSize(221, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.norden_sf.setFont(font)
        self.norden_sf.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.norden_sf.setText("")
        self.norden_sf.setAlignment(QtCore.Qt.AlignCenter)
        self.norden_sf.setObjectName("norden_sf")
        self.gridLayout.addWidget(self.norden_sf, 19, 2, 1, 1)
        self.tabla_base = QtWidgets.QTableWidget(self.scrollAreaWidgetContents)
        self.tabla_base.setMinimumSize(QtCore.QSize(1031, 280))
        self.tabla_base.setMaximumSize(QtCore.QSize(16777215, 280))
        self.tabla_base.setEditTriggers(QtWidgets.QAbstractItemView.AnyKeyPressed|QtWidgets.QAbstractItemView.DoubleClicked|QtWidgets.QAbstractItemView.EditKeyPressed|QtWidgets.QAbstractItemView.SelectedClicked)
        self.tabla_base.setObjectName("tabla_base")
        self.tabla_base.setColumnCount(0)
        self.tabla_base.setRowCount(0)
        self.gridLayout.addWidget(self.tabla_base, 1, 0, 7, 1)
        self.label = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label.setMinimumSize(QtCore.QSize(111, 16))
        self.label.setMaximumSize(QtCore.QSize(111, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)
        self.label_4 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_4.setMinimumSize(QtCore.QSize(71, 31))
        self.label_4.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_4.setFont(font)
        self.label_4.setStyleSheet("")
        self.label_4.setAlignment(QtCore.Qt.AlignCenter)
        self.label_4.setObjectName("label_4")
        self.gridLayout.addWidget(self.label_4, 1, 1, 1, 1)
        self.norden_base = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.norden_base.setMinimumSize(QtCore.QSize(221, 31))
        self.norden_base.setMaximumSize(QtCore.QSize(221, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.norden_base.setFont(font)
        self.norden_base.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.norden_base.setText("")
        self.norden_base.setAlignment(QtCore.Qt.AlignCenter)
        self.norden_base.setObjectName("norden_base")
        self.gridLayout.addWidget(self.norden_base, 1, 2, 1, 1)
        self.exportar_base = QtWidgets.QPushButton(self.scrollAreaWidgetContents)
        self.exportar_base.setMinimumSize(QtCore.QSize(93, 28))
        self.exportar_base.setMaximumSize(QtCore.QSize(93, 28))
        self.exportar_base.setObjectName("exportar_base")
        self.gridLayout.addWidget(self.exportar_base, 1, 3, 1, 1)
        self.label_8 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_8.setMinimumSize(QtCore.QSize(71, 31))
        self.label_8.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_8.setFont(font)
        self.label_8.setAlignment(QtCore.Qt.AlignCenter)
        self.label_8.setObjectName("label_8")
        self.gridLayout.addWidget(self.label_8, 5, 1, 1, 1)
        self.operador_base = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.operador_base.setMinimumSize(QtCore.QSize(221, 40))
        self.operador_base.setMaximumSize(QtCore.QSize(221, 40))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.operador_base.setFont(font)
        self.operador_base.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.operador_base.setText("")
        self.operador_base.setAlignment(QtCore.Qt.AlignCenter)
        self.operador_base.setObjectName("operador_base")
        self.gridLayout.addWidget(self.operador_base, 3, 2, 1, 1)
        self.label_7 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_7.setMinimumSize(QtCore.QSize(71, 31))
        self.label_7.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_7.setFont(font)
        self.label_7.setAlignment(QtCore.Qt.AlignCenter)
        self.label_7.setObjectName("label_7")
        self.gridLayout.addWidget(self.label_7, 4, 1, 1, 1)
        self.label_5 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_5.setMinimumSize(QtCore.QSize(71, 31))
        self.label_5.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_5.setFont(font)
        self.label_5.setAlignment(QtCore.Qt.AlignCenter)
        self.label_5.setObjectName("label_5")
        self.gridLayout.addWidget(self.label_5, 2, 1, 1, 1)
        self.termino_base = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.termino_base.setMinimumSize(QtCore.QSize(221, 31))
        self.termino_base.setMaximumSize(QtCore.QSize(221, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.termino_base.setFont(font)
        self.termino_base.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.termino_base.setText("")
        self.termino_base.setAlignment(QtCore.Qt.AlignCenter)
        self.termino_base.setObjectName("termino_base")
        self.gridLayout.addWidget(self.termino_base, 4, 2, 1, 1)
        self.label_9 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_9.setMinimumSize(QtCore.QSize(71, 31))
        self.label_9.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_9.setFont(font)
        self.label_9.setAlignment(QtCore.Qt.AlignCenter)
        self.label_9.setObjectName("label_9")
        self.gridLayout.addWidget(self.label_9, 6, 1, 1, 1)
        self.label_6 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_6.setMinimumSize(QtCore.QSize(71, 31))
        self.label_6.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_6.setFont(font)
        self.label_6.setAlignment(QtCore.Qt.AlignCenter)
        self.label_6.setObjectName("label_6")
        self.gridLayout.addWidget(self.label_6, 3, 1, 1, 1)
        self.mrc_base = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.mrc_base.setMinimumSize(QtCore.QSize(221, 31))
        self.mrc_base.setMaximumSize(QtCore.QSize(221, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.mrc_base.setFont(font)
        self.mrc_base.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.mrc_base.setText("")
        self.mrc_base.setAlignment(QtCore.Qt.AlignCenter)
        self.mrc_base.setObjectName("mrc_base")
        self.gridLayout.addWidget(self.mrc_base, 5, 2, 1, 1)
        self.nrc_base = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.nrc_base.setMinimumSize(QtCore.QSize(221, 31))
        self.nrc_base.setMaximumSize(QtCore.QSize(221, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.nrc_base.setFont(font)
        self.nrc_base.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.nrc_base.setText("")
        self.nrc_base.setAlignment(QtCore.Qt.AlignCenter)
        self.nrc_base.setObjectName("nrc_base")
        self.gridLayout.addWidget(self.nrc_base, 6, 2, 1, 1)
        self.id_base = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.id_base.setMinimumSize(QtCore.QSize(221, 31))
        self.id_base.setMaximumSize(QtCore.QSize(221, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.id_base.setFont(font)
        self.id_base.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.id_base.setText("")
        self.id_base.setAlignment(QtCore.Qt.AlignCenter)
        self.id_base.setObjectName("id_base")
        self.gridLayout.addWidget(self.id_base, 2, 2, 1, 1)
        self.norden_fact = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.norden_fact.setMinimumSize(QtCore.QSize(221, 31))
        self.norden_fact.setMaximumSize(QtCore.QSize(221, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.norden_fact.setFont(font)
        self.norden_fact.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.norden_fact.setText("")
        self.norden_fact.setAlignment(QtCore.Qt.AlignCenter)
        self.norden_fact.setObjectName("norden_fact")
        self.gridLayout.addWidget(self.norden_fact, 9, 2, 1, 1)
        self.label_36 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_36.setMinimumSize(QtCore.QSize(71, 31))
        self.label_36.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_36.setFont(font)
        self.label_36.setAlignment(QtCore.Qt.AlignCenter)
        self.label_36.setObjectName("label_36")
        self.gridLayout.addWidget(self.label_36, 24, 1, 1, 1)
        self.nrc_sf = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.nrc_sf.setMinimumSize(QtCore.QSize(221, 31))
        self.nrc_sf.setMaximumSize(QtCore.QSize(221, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.nrc_sf.setFont(font)
        self.nrc_sf.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.nrc_sf.setText("")
        self.nrc_sf.setAlignment(QtCore.Qt.AlignCenter)
        self.nrc_sf.setObjectName("nrc_sf")
        self.gridLayout.addWidget(self.nrc_sf, 24, 2, 1, 1)
        self.exportar_fact = QtWidgets.QPushButton(self.scrollAreaWidgetContents)
        self.exportar_fact.setMinimumSize(QtCore.QSize(93, 28))
        self.exportar_fact.setMaximumSize(QtCore.QSize(93, 28))
        self.exportar_fact.setObjectName("exportar_fact")
        self.gridLayout.addWidget(self.exportar_fact, 9, 3, 1, 1)
        self.label_2 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_2.setMinimumSize(QtCore.QSize(151, 16))
        self.label_2.setMaximumSize(QtCore.QSize(151, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 8, 0, 1, 1)
        self.nrc_fact = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.nrc_fact.setMinimumSize(QtCore.QSize(221, 31))
        self.nrc_fact.setMaximumSize(QtCore.QSize(221, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.nrc_fact.setFont(font)
        self.nrc_fact.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.nrc_fact.setText("")
        self.nrc_fact.setAlignment(QtCore.Qt.AlignCenter)
        self.nrc_fact.setObjectName("nrc_fact")
        self.gridLayout.addWidget(self.nrc_fact, 14, 2, 1, 1)
        self.tabla_fact = QtWidgets.QTableWidget(self.scrollAreaWidgetContents)
        self.tabla_fact.setMinimumSize(QtCore.QSize(1031, 280))
        self.tabla_fact.setMaximumSize(QtCore.QSize(16777215, 280))
        self.tabla_fact.setEditTriggers(QtWidgets.QAbstractItemView.AnyKeyPressed|QtWidgets.QAbstractItemView.DoubleClicked|QtWidgets.QAbstractItemView.EditKeyPressed|QtWidgets.QAbstractItemView.SelectedClicked)
        self.tabla_fact.setObjectName("tabla_fact")
        self.tabla_fact.setColumnCount(0)
        self.tabla_fact.setRowCount(0)
        self.gridLayout.addWidget(self.tabla_fact, 9, 0, 9, 1)
        self.label_17 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_17.setMinimumSize(QtCore.QSize(71, 31))
        self.label_17.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_17.setFont(font)
        self.label_17.setAlignment(QtCore.Qt.AlignCenter)
        self.label_17.setObjectName("label_17")
        self.gridLayout.addWidget(self.label_17, 14, 1, 1, 1)
        self.label_21 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_21.setMinimumSize(QtCore.QSize(71, 31))
        self.label_21.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_21.setFont(font)
        self.label_21.setAlignment(QtCore.Qt.AlignCenter)
        self.label_21.setObjectName("label_21")
        self.gridLayout.addWidget(self.label_21, 10, 1, 1, 1)
        self.label_32 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_32.setMinimumSize(QtCore.QSize(71, 31))
        self.label_32.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_32.setFont(font)
        self.label_32.setAlignment(QtCore.Qt.AlignCenter)
        self.label_32.setObjectName("label_32")
        self.gridLayout.addWidget(self.label_32, 23, 1, 1, 1)
        self.id_fact = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.id_fact.setMinimumSize(QtCore.QSize(221, 31))
        self.id_fact.setMaximumSize(QtCore.QSize(221, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.id_fact.setFont(font)
        self.id_fact.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.id_fact.setText("")
        self.id_fact.setAlignment(QtCore.Qt.AlignCenter)
        self.id_fact.setObjectName("id_fact")
        self.gridLayout.addWidget(self.id_fact, 10, 2, 1, 1)
        self.label_23 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_23.setMinimumSize(QtCore.QSize(71, 31))
        self.label_23.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_23.setFont(font)
        self.label_23.setAlignment(QtCore.Qt.AlignCenter)
        self.label_23.setObjectName("label_23")
        self.gridLayout.addWidget(self.label_23, 13, 1, 1, 1)
        self.operador_fact = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.operador_fact.setMinimumSize(QtCore.QSize(221, 40))
        self.operador_fact.setMaximumSize(QtCore.QSize(221, 40))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.operador_fact.setFont(font)
        self.operador_fact.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.operador_fact.setText("")
        self.operador_fact.setAlignment(QtCore.Qt.AlignCenter)
        self.operador_fact.setObjectName("operador_fact")
        self.gridLayout.addWidget(self.operador_fact, 11, 2, 1, 1)
        self.label_20 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_20.setMinimumSize(QtCore.QSize(71, 31))
        self.label_20.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_20.setFont(font)
        self.label_20.setAlignment(QtCore.Qt.AlignCenter)
        self.label_20.setObjectName("label_20")
        self.gridLayout.addWidget(self.label_20, 11, 1, 1, 1)
        self.label_25 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_25.setMinimumSize(QtCore.QSize(71, 31))
        self.label_25.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_25.setFont(font)
        self.label_25.setStyleSheet("")
        self.label_25.setAlignment(QtCore.Qt.AlignCenter)
        self.label_25.setObjectName("label_25")
        self.gridLayout.addWidget(self.label_25, 9, 1, 1, 1)
        self.label_19 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_19.setMinimumSize(QtCore.QSize(71, 31))
        self.label_19.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_19.setFont(font)
        self.label_19.setAlignment(QtCore.Qt.AlignCenter)
        self.label_19.setObjectName("label_19")
        self.gridLayout.addWidget(self.label_19, 12, 1, 1, 1)
        self.termino_sf = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.termino_sf.setMinimumSize(QtCore.QSize(221, 31))
        self.termino_sf.setMaximumSize(QtCore.QSize(221, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.termino_sf.setFont(font)
        self.termino_sf.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.termino_sf.setText("")
        self.termino_sf.setAlignment(QtCore.Qt.AlignCenter)
        self.termino_sf.setObjectName("termino_sf")
        self.gridLayout.addWidget(self.termino_sf, 22, 2, 1, 1)
        self.termino_fact = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.termino_fact.setMinimumSize(QtCore.QSize(221, 31))
        self.termino_fact.setMaximumSize(QtCore.QSize(221, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.termino_fact.setFont(font)
        self.termino_fact.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.termino_fact.setText("")
        self.termino_fact.setAlignment(QtCore.Qt.AlignCenter)
        self.termino_fact.setObjectName("termino_fact")
        self.gridLayout.addWidget(self.termino_fact, 12, 2, 1, 1)
        self.mrc_fact = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.mrc_fact.setMinimumSize(QtCore.QSize(221, 31))
        self.mrc_fact.setMaximumSize(QtCore.QSize(221, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.mrc_fact.setFont(font)
        self.mrc_fact.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.mrc_fact.setText("")
        self.mrc_fact.setAlignment(QtCore.Qt.AlignCenter)
        self.mrc_fact.setObjectName("mrc_fact")
        self.gridLayout.addWidget(self.mrc_fact, 13, 2, 1, 1)
        self.label_31 = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.label_31.setMinimumSize(QtCore.QSize(71, 31))
        self.label_31.setMaximumSize(QtCore.QSize(71, 31))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.label_31.setFont(font)
        self.label_31.setAlignment(QtCore.Qt.AlignCenter)
        self.label_31.setObjectName("label_31")
        self.gridLayout.addWidget(self.label_31, 22, 1, 1, 1)
        self.groupBox = QtWidgets.QGroupBox(self.scrollAreaWidgetContents)
        self.groupBox.setEnabled(True)
        self.groupBox.setMinimumSize(QtCore.QSize(393, 28))
        self.groupBox.setMaximumSize(QtCore.QSize(393, 28))
        self.groupBox.setSizeIncrement(QtCore.QSize(0, 0))
        self.groupBox.setMouseTracking(False)
        self.groupBox.setAutoFillBackground(False)
        self.groupBox.setTitle("")
        self.groupBox.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignTop)
        self.groupBox.setFlat(False)
        self.groupBox.setCheckable(False)
        self.groupBox.setObjectName("groupBox")
        self.buscar_button = QtWidgets.QPushButton(self.groupBox)
        self.buscar_button.setGeometry(QtCore.QRect(0, 0, 93, 28))
        self.buscar_button.setMinimumSize(QtCore.QSize(93, 28))
        self.buscar_button.setMaximumSize(QtCore.QSize(93, 28))
        self.buscar_button.setObjectName("buscar_button")
        self.agregar_button = QtWidgets.QPushButton(self.groupBox)
        self.agregar_button.setGeometry(QtCore.QRect(100, 0, 93, 28))
        self.agregar_button.setMinimumSize(QtCore.QSize(93, 28))
        self.agregar_button.setMaximumSize(QtCore.QSize(93, 28))
        self.agregar_button.setObjectName("agregar_button")
        self.modificar_button = QtWidgets.QPushButton(self.groupBox)
        self.modificar_button.setGeometry(QtCore.QRect(200, 0, 93, 28))
        self.modificar_button.setMinimumSize(QtCore.QSize(93, 28))
        self.modificar_button.setMaximumSize(QtCore.QSize(93, 28))
        self.modificar_button.setObjectName("modificar_button")
        self.actualizar_button = QtWidgets.QPushButton(self.groupBox)
        self.actualizar_button.setGeometry(QtCore.QRect(300, 0, 93, 28))
        self.actualizar_button.setMinimumSize(QtCore.QSize(93, 28))
        self.actualizar_button.setMaximumSize(QtCore.QSize(93, 28))
        self.actualizar_button.setObjectName("actualizar_button")
        self.gridLayout.addWidget(self.groupBox, 26, 1, 1, 2)
        self.tabla_base.raise_()
        self.label_3.raise_()
        self.label_2.raise_()
        self.tabla_fact.raise_()
        self.label.raise_()
        self.tabla_sf.raise_()
        self.label_3.raise_()
        self.label_2.raise_()
        self.tabla_fact.raise_()
        self.tabla_sf.raise_()
        self.label_3.raise_()
        self.label_2.raise_()
        self.tabla_fact.raise_()
        self.tabla_sf.raise_()
        self.norden_base.raise_()
        self.label_4.raise_()
        self.norden_sf.raise_()
        self.label_35.raise_()
        self.groupBox.raise_()
        self.id_base.raise_()
        self.label_5.raise_()
        self.operador_base.raise_()
        self.label_6.raise_()
        self.termino_base.raise_()
        self.label_7.raise_()
        self.mrc_base.raise_()
        self.label_8.raise_()
        self.nrc_base.raise_()
        self.label_9.raise_()
        self.exportar_base.raise_()
        self.exportar_sf.raise_()
        self.id_sf.raise_()
        self.label_29.raise_()
        self.operador_sf.raise_()
        self.label_33.raise_()
        self.termino_sf.raise_()
        self.label_31.raise_()
        self.mrc_sf.raise_()
        self.label_32.raise_()
        self.nrc_sf.raise_()
        self.label_36.raise_()
        self.select_sf.raise_()
        self.norden_fact.raise_()
        self.label_25.raise_()
        self.id_fact.raise_()
        self.label_21.raise_()
        self.operador_fact.raise_()
        self.label_20.raise_()
        self.termino_fact.raise_()
        self.label_19.raise_()
        self.mrc_fact.raise_()
        self.label_23.raise_()
        self.nrc_fact.raise_()
        self.label_17.raise_()
        self.exportar_fact.raise_()
        self.scrollArea.setWidget(self.scrollAreaWidgetContents)
        self.gridLayout_2.addWidget(self.scrollArea, 1, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1580, 26))
        self.menubar.setObjectName("menubar")
        self.menuMenu = QtWidgets.QMenu(self.menubar)
        self.menuMenu.setObjectName("menuMenu")
        self.menuCargar = QtWidgets.QMenu(self.menubar)
        self.menuCargar.setObjectName("menuCargar")        
        self.menuSettings = QtWidgets.QMenu(self.menubar)
        self.menuSettings.setObjectName("menuSettings")       
        self.menuRegion = QtWidgets.QMenu(self.menuSettings)
        self.menuRegion.setObjectName("menuRegion")
        MainWindow.setMenuBar(self.menubar)       
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionCargar_Archivo_Base = QtWidgets.QAction(MainWindow)
        self.actionCargar_Archivo_Base.setObjectName("actionCargar_Archivo_Base")
        self.actionCargar_Archivo_Facturacion = QtWidgets.QAction(MainWindow)
        self.actionCargar_Archivo_Facturacion.setObjectName("actionCargar_Archivo_Facturacion")
        self.actionCargar_Archivo_Salesforce = QtWidgets.QAction(MainWindow)
        self.actionCargar_Archivo_Salesforce.setObjectName("actionCargar_Archivo_Salesforce")
        self.actionCambiar_Usuario = QtWidgets.QAction(MainWindow)
        self.actionCambiar_Usuario.setObjectName("actionCambiar_Usuario")
        self.actionExportar_todo = QtWidgets.QAction(MainWindow)
        self.actionExportar_todo.setObjectName("actionExportar_todo")
        self.actionSalir = QtWidgets.QAction(MainWindow)
        self.actionSalir.setObjectName("actionSalir")
        self.actionManual = QtWidgets.QAction(MainWindow)
        self.actionManual.setObjectName("actionManual")
        self.actionSeleccionar_Fecha = QtWidgets.QAction(MainWindow)
        self.actionSeleccionar_Fecha.setCheckable(True)
        self.actionSeleccionar_Fecha.setObjectName("actionSeleccionar_Fecha")
        self.actionRegional = QtWidgets.QAction(MainWindow)
        self.actionRegional.setCheckable(True)
        self.actionRegional.setObjectName("actionRegional")
        self.actionColombia = QtWidgets.QAction(MainWindow)
        self.actionColombia.setCheckable(True)
        self.actionColombia.setObjectName("actionColombia")
        self.menuMenu.addAction(self.actionCambiar_Usuario)
        self.menuMenu.addAction(self.actionExportar_todo)
        self.menuMenu.addAction(self.actionManual)
        self.menuMenu.addAction(self.actionSalir)
        self.menuCargar.addAction(self.actionCargar_Archivo_Base)
        self.menuCargar.addAction(self.actionCargar_Archivo_Facturacion)
        self.menuCargar.addAction(self.actionCargar_Archivo_Salesforce)
        self.menubar.addAction(self.menuMenu.menuAction())
        self.menubar.addAction(self.menuCargar.menuAction())
        self.menuRegion.addAction(self.actionRegional)
        self.menuRegion.addAction(self.actionColombia)
        self.menuSettings.addAction(self.actionSeleccionar_Fecha)
        self.menuSettings.addAction(self.menuRegion.menuAction())
        self.menubar.addAction(self.menuSettings.menuAction())
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        self.buscartodo_button = QtWidgets.QPushButton(self.scrollAreaWidgetContents)
        self.buscartodo_button.setMinimumSize(QtCore.QSize(93, 28))
        self.buscartodo_button.setMaximumSize(QtCore.QSize(93, 28))
        self.buscartodo_button.setText('Buscar Todo')
        self.gridLayout.addWidget(self.buscartodo_button, 26, 3, 1, 2)
        """
# =============================================================================
#         END
# =============================================================================
        """
        
        import RAP_rc
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/Register/logo2.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setWindowState(MainWindow.windowState() & ~QtCore.Qt.WindowMinimized | QtCore.Qt.WindowActive)
        MainWindow.setFocus(QtCore.Qt.PopupFocusReason)
        MainWindow.raise_()
        self.actionCargar_Archivo_Base.triggered.connect(self.getxlsbase)
        self.actionCargar_Archivo_Facturacion.triggered.connect(self.getxlsfacturacion)
        self.actionCargar_Archivo_Salesforce.triggered.connect(self.getxlssf)
        self.actionSeleccionar_Fecha.triggered.connect(self.openPopUp)
        self.actualizar_button.clicked.connect(self.actualizarsf)
        self.tabla_sf.cellClicked.connect(self.cellselect)
        self.select_sf.currentIndexChanged.connect(self.selectsf)
        self.exportar_base.clicked.connect(self.exp_base)
        self.buscartodo_button.clicked.connect(self.buscar_todo)
        self.agregar_button.clicked.connect(self.agregar_base)


    """     
# =============================================================================
#  SE CARGAN LOS ARCHIVOS  & SE FILTRA POR FECHA
# =============================================================================
    """
    
    def actualizarsf(self):
        self.select_sf.setCurrentIndex(0)
        self.filtered_region_index=list()
        self.filtered_colombia_index=list()
        num_row = self.hoja_sf.nrows-5
        num_col = self.hoja_sf.ncols
        if 'date_select' in globals():
            if self.actionSeleccionar_Fecha.isChecked():
                monthsf = date_select.month
                yearsf = date_select.year
                self.statusbar.showMessage('Periodo seleccionado: %i/%i' %(date_select.month, date_select.year))
                self.filtered_date_index=list()
                self.filtered_dr_index=list()
                self.filtered_dc_index=list()
                self.filtered_region_index=list()
                num_row = self.hoja_sf.nrows-5
                num_col = self.hoja_sf.ncols
                for i in range(1,num_row):
                    for j in range(num_col):
                        if j == 10 :
                            valor10 =self.datasf[i][j]
                            y = type(valor10) is float
                            if y == True:
                                seconds10 = (valor10 - 25569) * 86400.0
                                xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                xty= datetime.datetime.utcfromtimestamp(seconds10).year
                                if xtm == monthsf:
                                    if xty == yearsf:
                                        self.filtered_date_index.append(i)
                                        self.filtered_date_index.sort()


                if self.actionRegional.isChecked(): #despliega info filtrada por fecha y regional
                    for i in self.filtered_date_index:
                        for j in range(num_col):
                            valor = self.datasf[i][j]
                            if j == 25:
                                if valor == 'Wholesale Regional':
                                    self.filtered_dr_index.append(i)
                                    self.filtered_dr_index.sort()
                    self.tabla_sf.setRowCount(len(self.filtered_dr_index))
                    l=0
                    for i in self.filtered_dr_index:
                        
                        for j in range(num_col):
                            if i == 0:
                                self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(i,j).value)))
                            else: 
                                if j == 10 or j == 2 or j == 9:
                                    valor10 =self.datasf[i][j]
                                    y = type(valor10) is float
                                    if y == True:
                                        if valor10 == 1:
                                            nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                        else:
                                            seconds10 = (valor10 - 25569) * 86400.0
                                            xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                            xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                            nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                    else:
                                        valor = str(self.datasf[i][j])
                                        nvalor = QtWidgets.QTableWidgetItem(valor)
                                        self.tabla_sf.setItem(l,j, nvalor)
                                else:
                                    
                                    valor = str(self.datasf[i][j])
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_sf.setItem(l,j, nvalor)
                            self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(0,j).value)))
                        l=l+1
                        
                elif self.actionColombia.isChecked(): #despliega info filtrada por fecha y colombia
                    for i in self.filtered_date_index:
                        for j in range(num_col):
                            valor = self.datasf[i][j]
                            if j== 29:
                                if valor == 'CN Local Colombia':
                                    self.filtered_dc_index.append(i)
                                    self.filtered_dc_index.sort()
                    
                    self.tabla_sf.setRowCount(len(self.filtered_dc_index))
                    l=0
                    for i in self.filtered_dc_index:
                        
                        for j in range(num_col):
                            if i == 0:
                                self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(i,j).value)))
                            else: 
                                if j == 10 or j == 2 or j == 9:
                                    valor10 =self.datasf[i][j]
                                    y = type(valor10) is float
                                    if y == True:
                                        if valor10 == 1:
                                            nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                        else:
                                            seconds10 = (valor10 - 25569) * 86400.0
                                            xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                            xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                            nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                    else:
                                        valor = str(self.datasf[i][j])
                                        nvalor = QtWidgets.QTableWidgetItem(valor)
                                        self.tabla_sf.setItem(l,j, nvalor)
                                else:
                                    
                                    valor = str(self.datasf[i][j])
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_sf.setItem(l,j, nvalor)
                            self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(0,j).value)))
                        l=l+1
                else:  #despliega info filtrada por fecha   
                    
                    self.tabla_sf.setRowCount(len(self.filtered_date_index))
                    l=0
                    for i in self.filtered_date_index:
                        
                        for j in range(num_col):
                            if i == 0:
                                self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(i,j).value)))
                            else: 
                                if j == 10 or j == 2 or j == 9:
                                    valor10 =self.datasf[i][j]
                                    y = type(valor10) is float
                                    if y == True:
                                        if valor10 == 1:
                                            nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                        else:
                                            seconds10 = (valor10 - 25569) * 86400.0
                                            xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                            xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                            nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                    else:
                                        valor = str(self.datasf[i][j])
                                        nvalor = QtWidgets.QTableWidgetItem(valor)
                                        self.tabla_sf.setItem(l,j, nvalor)
                                else:
                                    
                                    valor = str(self.datasf[i][j])
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_sf.setItem(l,j, nvalor)
                            self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(0,j).value)))
                        l=l+1
                    
                    
                monthf = date_select.month
                yearf = date_select.year
                if monthf == 12:
                    monthf = 1
                    yearf = date_select.year + 1
                else:
                    monthf = date_select.month + 1
                self.filtered_date_index_fact=list()
                for i in range(1,self.hoja_facturacion.nrows):
                    for j in range(self.hoja_facturacion.ncols):
                        if j == 11 :
                            valor10 =self.dataf[i][j]
                            y = type(valor10) is float
                            if y == True and valor10 != 1:
                                seconds10 = (valor10 - 25569) * 86400.0
                                xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                xty= datetime.datetime.utcfromtimestamp(seconds10).year
                                xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                if xtm == monthf:
                                    if xty == yearf:
                                        self.filtered_date_index_fact.append(i)
                                        self.filtered_date_index_fact.sort()
            
                self.tabla_fact.setRowCount(len(self.filtered_date_index_fact))
                lf=0
                for i in self.filtered_date_index_fact:
                    for j in range(num_col):
                        self.tabla_fact.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_facturacion.cell(0,j).value)))
                        if i == 0:
                            self.tabla_fact.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_facturacion.cell(i,j).value)))
                        else: 
                            if j==0 or j == 10 or j == 11 or j == 14:
                                valor10 =self.dataf[i][j]
                                y = type(valor10) is float
                                if y == True:
                                    if valor10 == 1:
                                        nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                        self.tabla_fact.setItem(lf,j, nvalor10)
                                    else:
                                        seconds10 = (valor10 - 25569) * 86400.0
                                        xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                        xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                        nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                        self.tabla_fact.setItem(lf,j, nvalor10)
                                else:
                                    valor = str(self.dataf[i][j])
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_fact.setItem(lf,j, nvalor)
                            else:
                                
                                valor = str(self.dataf[i][j])
                                nvalor = QtWidgets.QTableWidgetItem(valor)
                                self.tabla_fact.setItem(lf,j, nvalor)
                    self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_facturacion.cell(i,j).value)))
                    lf=lf+1         
            else: #no está seleccionada la fecha
                if self.actionRegional.isChecked(): #despliega info filtrada por regional
                    for i in range(num_row):
                        for j in range(num_col):
                            valor = self.datasf[i][j]
                            if j == 25:
                                if valor == 'Wholesale Regional':
                                    self.filtered_region_index.append(i)
                                    self.filtered_region_index.sort()
                    self.tabla_sf.setRowCount(len(self.filtered_region_index))
                    l=0
                    for i in self.filtered_region_index:
                        
                        for j in range(num_col):
                            if i == 0:
                                self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(i,j).value)))
                            else: 
                                if j == 10 or j == 2 or j == 9:
                                    valor10 =self.datasf[i][j]
                                    y = type(valor10) is float
                                    if y == True:
                                        if valor10 == 1:
                                            nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                        else:
                                            seconds10 = (valor10 - 25569) * 86400.0
                                            xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                            xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                            nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                    else:
                                        valor = str(self.datasf[i][j])
                                        nvalor = QtWidgets.QTableWidgetItem(valor)
                                        self.tabla_sf.setItem(l,j, nvalor)
                                else:
                                    
                                    valor = str(self.datasf[i][j])
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_sf.setItem(l,j, nvalor)
                            self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(0,j).value)))
                        l=l+1
                        
                        
                elif self.actionColombia.isChecked(): #despliega info filtrada colombia
                    for i in self.filtered_date_index:
                        for j in range(num_col):
                            valor = self.datasf[i][j]
                            if j== 29:
                                if valor == 'CN Local Colombia':
                                    self.filtered_colombia_index.append(i)
                                    self.filtered_colombia_index.sort()
                    self.tabla_sf.setRowCount(len(self.filtered_colombia_index))
                    l=0
                    for i in self.filtered_colombia_index:
                        
                        for j in range(num_col):
                            if i == 0:
                                self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(i,j).value)))
                            else: 
                                if j == 10 or j == 2 or j == 9:
                                    valor10 =self.datasf[i][j]
                                    y = type(valor10) is float
                                    if y == True:
                                        if valor10 == 1:
                                            nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                        else:
                                            seconds10 = (valor10 - 25569) * 86400.0
                                            xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                            xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                            nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                    else:
                                        valor = str(self.datasf[i][j])
                                        nvalor = QtWidgets.QTableWidgetItem(valor)
                                        self.tabla_sf.setItem(l,j, nvalor)
                                else:
                                    
                                    valor = str(self.datasf[i][j])
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_sf.setItem(l,j, nvalor)
                            self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(0,j).value)))
                        l=l+1                    
                else:  #despliega info completa 
                    self.tabla_sf.setRowCount(len(self.datasf)-6)
                    for i in range(len(self.datasf)):
                        for j in range(num_col):
                            if i == 0:
                                self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(i,j).value)))
                            else: 
                                if j == 10 or j == 2 or j == 9:
                                    valor10 =self.hoja_sf.cell(i,j).value
                                    y = type(valor10) is float
                                    if y == True:
                                        seconds10 = (valor10 - 25569) * 86400.0
                                        xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                        xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                        nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                        self.tabla_sf.setItem(i-1,j, nvalor10)
                                    else:
                                        valor = str(self.hoja_sf.cell(i,j).value)
                                        nvalor = QtWidgets.QTableWidgetItem(valor)
                                        self.tabla_sf.setItem(i-1,j, nvalor)
                                else:
                                    
                                    valor = str(self.hoja_sf.cell(i,j).value)
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_sf.setItem(i-1,j, nvalor)
                    



            
                self.tabla_fact.setRowCount(len(self.dataf)-1)
                for i in range(len(self.dataf)):
                    for j in range(self.hoja_facturacion.ncols):
                        if i == 0:
                            self.tabla_fact.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_facturacion.cell(i,j).value)))
                        else: 
                            if j == 10 or j == 2 or j == 9:
                                valor10 =self.hoja_facturacion.cell(i,j).value
                                y = type(valor10) is float
                                if y == True:
                                    if valor10 >1:
                                        seconds10 = (valor10 - 25569) * 86400.0
                                        xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                        xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                        nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                        self.tabla_fact.setItem(i-1,j, nvalor10)
                                    else:
                                        nvalor = QtWidgets.QTableWidgetItem('1/01/1900')
                                        self.tabla_fact.setItem(i-1,j, nvalor)
                                else:
                                    valor = str(self.hoja_facturacion.cell(i,j).value)
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_fact.setItem(i-1,j, nvalor)
                            else:
                                
                                valor = str(self.hoja_facturacion.cell(i,j).value)
                                nvalor = QtWidgets.QTableWidgetItem(valor)
                                self.tabla_fact.setItem(i-1,j, nvalor)
                            
            self.tabla_fact.resizeColumnsToContents()
                    
        else: #filtra por region (copiar el bloque anterior)    
            self.filtered_region_index=list()
            self.filtered_colombia_index=list()
            num_row = len(self.datasf)-6
            num_col = self.hoja_sf.ncols
            if self.actionRegional.isChecked():
                for i in range(num_row):
                    for j in range(num_col):
                     #despliega info filtrada por regional
                        valor = self.datasf[i][j]
                        if j == 25:
                            if valor == 'Wholesale Regional':
                                self.filtered_region_index.append(i)
                                self.filtered_region_index.sort()
            
                self.tabla_sf.setRowCount(len(self.filtered_region_index))
                l=0
                for i in self.filtered_region_index:
                    for j in range(num_col):
                        if i == 0:
                            self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(i,j).value)))
                        else: 
                            if j == 10 or j == 2 or j == 9:
                                valor10 =self.datasf[i][j]
                                y = type(valor10) is float
                                if y == True:
                                    if valor10 == 1:
                                        nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                        self.tabla_sf.setItem(l,j, nvalor10)
                                    else:
                                        seconds10 = (valor10 - 25569) * 86400.0
                                        xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                        xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                        nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                        self.tabla_sf.setItem(l,j, nvalor10)
                                else:
                                    valor = str(self.datasf[i][j])
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_sf.setItem(l,j, nvalor)
                            else:
                                
                                valor = str(self.datasf[i][j])
                                nvalor = QtWidgets.QTableWidgetItem(valor)
                                self.tabla_sf.setItem(l,j, nvalor)
                        self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(0,j).value)))
                    l=l+1
                        
                        
            elif self.actionColombia.isChecked(): #despliega info filtrada colombia
                
                for i in range(num_row):
                    for j in range(num_col):
                     #despliega info filtrada por regional
                        valor = self.datasf[i][j]
                        if j == 29:
                            if valor == 'CN Local Colombia':
                                self.filtered_colombia_index.append(i)
                                self.filtered_colombia_index.sort()
    
                self.tabla_sf.setRowCount(len(self.filtered_colombia_index))
                l=0
                for i in self.filtered_colombia_index:
                    for j in range(num_col):
                        if i == 0:
                            self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(i,j).value)))
                        else: 
                            if j == 10 or j == 2 or j == 9:
                                valor10 =self.datasf[i][j]
                                y = type(valor10) is float
                                if y == True:
                                    if valor10 == 1:
                                        nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                        self.tabla_sf.setItem(l,j, nvalor10)
                                    else:
                                        seconds10 = (valor10 - 25569) * 86400.0
                                        xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                        xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                        nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                        self.tabla_sf.setItem(l,j, nvalor10)
                                else:
                                    valor = str(self.datasf[i][j])
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_sf.setItem(l,j, nvalor)
                            else:
                                
                                valor = str(self.datasf[i][j])
                                nvalor = QtWidgets.QTableWidgetItem(valor)
                                self.tabla_sf.setItem(l,j, nvalor)
                        self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(0,j).value)))
                    l=l+1                    
            else:  #despliega info completa
                self.tabla_sf.setRowCount(len(self.datasf)-7)
                print(len(self.datasf))
                for i in range(len(self.datasf)):
                    for j in range(num_col):
                        if i == 0:
                            self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(0,j).value)))
                        else: 
                            if j == 10 or j == 2 or j == 9:
                                valor10 = self.datasf[i][j]
                                y = type(valor10) is float
                                if y == True:
                                    seconds10 = (valor10 - 25569) * 86400.0
                                    xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                    xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                    nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                    self.tabla_sf.setItem(i-1,j, nvalor10)
                                else:
                                    valor = str(self.datasf[i][j])
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_sf.setItem(i-1,j, nvalor)
                            else:
                                
                                valor = str(self.datasf[i][j])
                                nvalor = QtWidgets.QTableWidgetItem(valor)
                                self.tabla_sf.setItem(i-1,j, nvalor)
            self.tabla_sf.resizeColumnsToContents()
            

            
        self.tabla_busquedasf = np.empty((self.tabla_sf.rowCount(), self.tabla_sf.columnCount()), dtype=('U100'))
        for x in range(self.tabla_sf.rowCount()):
               for y in range(self.tabla_sf.columnCount()):
                   self.tabla_busquedasf[x][y] = self.tabla_sf.item(x,y).text()    
        self.actionSeleccionar_Fecha.setChecked(False)
        self.actionRegional.setChecked(False)
        self.actionColombia.setChecked(False)
        self.tabla_sf.resizeColumnsToContents() 
    def getxlsbase(self):
        if 'date_select' in globals():
            self.statusbar.showMessage('Periodo seleccionado: %i/%i' %(date_select.month, date_select.year))
        filePath_base, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'Select file', './', 'Excel Files (*.xls *.xlsx)')
        base = xlrd.open_workbook(filePath_base)
        self.hoja_base = base.sheet_by_index(0)
        self.datab = [[self.hoja_base.cell_value(r, c) for c in range(self.hoja_base.ncols)] for r in range(self.hoja_base.nrows)]
        num_row = self.hoja_base.nrows
        num_col = self.hoja_base.ncols 
        self.tabla_base.setRowCount(num_row-3)
        self.tabla_base.setColumnCount(num_col)
        self.header = self.tabla_base.horizontalHeader()
        self.tabla_base.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        for i in range(num_row):
            for j in range(num_col):
                if i == 2:
                    self.tabla_base.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_base.cell(i,j).value)))
                elif i>2:
                    if j == 15 or j == 17:
                        valor10 =self.hoja_base.cell(i,j).value
                        if valor10 == '':
                            valor = str(self.hoja_base.cell(i,j).value)
                            nvalor = QtWidgets.QTableWidgetItem(valor)
                            self.tabla_base.setItem(i-3,j, nvalor)
                        else: 
                            valor10 = float(valor10)
                            y = type(valor10) is float
                            if y == True:
                                if valor10 == 1:
                                    nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                    self.tabla_base.setItem(i-3,j, nvalor10)
                                else:
                                    seconds10 = (valor10 - 25569) * 86400.0
                                    xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                    nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                    self.tabla_base.setItem(i-3,j, nvalor10)
                            else:
                                valor = str(self.hoja_base.cell(i,j).value)
                                nvalor = QtWidgets.QTableWidgetItem(valor)
                                self.tabla_base.setItem(i-3,j, nvalor)
                    else:
                        valor = str(self.hoja_base.cell(i,j).value)
                        if valor == '42':
                            nvalor = QtWidgets.QTableWidgetItem('N/A')
                            self.tabla_base.setItem(i-3,j, nvalor)
                        else:
                            nvalor = QtWidgets.QTableWidgetItem(valor)
                            self.tabla_base.setItem(i-3,j, nvalor)
        self.tabla_base.resizeColumnsToContents()  
        
    def getxlsfacturacion(self):
        filePath_facturacion, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'Select file', './', 'Excel Files (*.xls *.xlsx)')
        facturacion = xlrd.open_workbook(filePath_facturacion)
        self.hoja_facturacion = facturacion.sheet_by_index(0)
        self.dataf = [[self.hoja_facturacion.cell_value(r, c) for c in range(self.hoja_facturacion.ncols)] for r in range(self.hoja_facturacion.nrows)]
        num_row = self.hoja_facturacion.nrows
        num_col = self.hoja_facturacion.ncols 
        self.tabla_fact.setRowCount(num_row-1)
        self.tabla_fact.setColumnCount(num_col)
        self.header = self.tabla_fact.horizontalHeader()
        self.tabla_fact.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        
        if 'date_select' in globals():
            self.statusbar.showMessage('Periodo seleccionado: %i/%i' %(date_select.month, date_select.year))
            monthf = date_select.month
            yearf = date_select.year
            if monthf == 12:
                monthf = 1
                yearf = date_select.year + 1
            else:
                monthf = date_select.month + 1
            self.filtered_date_index_fact=list()
            for i in range(1,num_row):
                for j in range(num_col):
                    if j == 11 :
                        valor10 =self.dataf[i][j]
                        y = type(valor10) is float
                        if y == True and valor10 != 1:
                            seconds10 = (valor10 - 25569) * 86400.0
                            xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                            xty= datetime.datetime.utcfromtimestamp(seconds10).year
                            xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                            if xtm == monthf:
                                if xty == yearf:
                                    self.filtered_date_index_fact.append(i)
                                    self.filtered_date_index_fact.sort()
        
        if self.actionSeleccionar_Fecha.isChecked():
            l=0
            for i in self.filtered_date_index_fact:
                
                for j in range(num_col):
                    self.tabla_fact.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_facturacion.cell(0,j).value)))
                    if i == 0:
                        self.tabla_fact.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_facturacion.cell(i,j).value)))
                    else: 
                        if j==0 or j == 10 or j == 11 or j == 14:
                            valor10 =self.hoja_facturacion.cell(i,j).value
                            y = type(valor10) is float
                            if y == True:
                                if valor10 == 1:
                                    nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                    self.tabla_fact.setItem(l,j, nvalor10)
                                else:
                                    seconds10 = (valor10 - 25569) * 86400.0
                                    xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                    xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                    nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                    self.tabla_fact.setItem(l,j, nvalor10)
                            else:
                                valor = str(self.hoja_facturacion.cell(i,j).value)
                                nvalor = QtWidgets.QTableWidgetItem(valor)
                                self.tabla_fact.setItem(l,j, nvalor)
                        else:
                            
                            valor = str(self.hoja_facturacion.cell(i,j).value)
                            nvalor = QtWidgets.QTableWidgetItem(valor)
                            self.tabla_fact.setItem(l,j, nvalor)
                l=l+1
            
            self.tabla_fact.setRowCount(len(self.filtered_date_index_fact))
        
        else:
            for i in range(num_row):
                for j in range(num_col):
                    if i == 0:
                        self.tabla_fact.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_facturacion.cell(i,j).value)))
                    elif i>0:
                        if j==0 or j == 10 or j == 11 or j == 14:
                            valor10 =self.hoja_facturacion.cell(i,j).value
                            if valor10 == '':
                                valor = str(self.hoja_facturacion.cell(i,j).value)
                                nvalor = QtWidgets.QTableWidgetItem(valor)
                                self.tabla_fact.setItem(i-1,j, nvalor)
                            else: 
                                valor10 = float(valor10)
                                y = type(valor10) is float
                                if y == True:
                                    if valor10 == 1:
                                        nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                        self.tabla_fact.setItem(i-1,j, nvalor10)
                                    else:
                                        seconds10 = (valor10 - 25569) * 86400.0
                                        xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                        nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                        self.tabla_fact.setItem(i-1,j, nvalor10)
                        else:
                            valor = str(self.hoja_facturacion.cell(i,j).value)
                            if valor == '42':
                                nvalor = QtWidgets.QTableWidgetItem('N/A')
                                self.tabla_fact.setItem(i-1,j, nvalor)
                            else:
                                nvalor = QtWidgets.QTableWidgetItem(valor)
                                self.tabla_fact.setItem(i-1,j, nvalor)
        self.tabla_fact.resizeColumnsToContents()




    def getxlssf(self):
        filePath_sf, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'Select file', './', 'Excel Files (*.xls *.xlsx)')
        sf = xlrd.open_workbook(filePath_sf)
        self.hoja_sf = sf.sheet_by_index(0)
        self.datasf = [[self.hoja_sf.cell_value(r, c) for c in range(self.hoja_sf.ncols)] for r in range(self.hoja_sf.nrows)]
        num_row = self.hoja_sf.nrows-5
        num_col = self.hoja_sf.ncols 
        self.tabla_sf.setRowCount(num_row-2)
        self.tabla_sf.setColumnCount(num_col)
        self.header = self.tabla_sf.horizontalHeader()
        self.tabla_sf.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)

        if 'date_select' in globals():
            monthsf = date_select.month
            yearsf = date_select.year
            self.statusbar.showMessage('Periodo seleccionado: %i/%i' %(date_select.month, date_select.year))
            self.filtered_date_index=list()
            self.filtered_dr_index=list()
            self.filtered_region_index=list()
            for i in range(1,num_row):
                for j in range(num_col):
                    if j == 10 :
                        valor10 =self.datasf[i][j]
                        y = type(valor10) is float
                        if y == True:
                            seconds10 = (valor10 - 25569) * 86400.0
                            xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                            xty= datetime.datetime.utcfromtimestamp(seconds10).year
                            if xtm == monthsf:
                                if xty == yearsf:
                                    self.filtered_date_index.append(i)
                                    self.filtered_date_index.sort()

            if self.actionSeleccionar_Fecha.isChecked():
                if self.actionRegional.isChecked(): #despliega info filtrada por fecha y regional
                    for i in self.filtered_date_index:
                        for j in range(num_col):
                            valor = self.datasf[i][j]
                            if j == 25:
                                if valor == 'Wholesale Regional':
                                    self.filtered_dr_index.append(i)
                                    self.filtered_dr_index.sort()
                    self.tabla_sf.setRowCount(len(self.filtered_dr_index))
                    l=0
                    for i in self.filtered_dr_index:
                        
                        for j in range(num_col):
                            if i == 0:
                                self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(i,j).value)))
                            else: 
                                if j == 10 or j == 2 or j == 9:
                                    valor10 =self.datasf[i][j]
                                    y = type(valor10) is float
                                    if y == True:
                                        if valor10 == 1:
                                            nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                        else:
                                            seconds10 = (valor10 - 25569) * 86400.0
                                            xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                            xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                            nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                    else:
                                        valor = str(self.datasf[i][j])
                                        nvalor = QtWidgets.QTableWidgetItem(valor)
                                        self.tabla_sf.setItem(l,j, nvalor)
                                else:
                                    
                                    valor = str(self.datasf[i][j])
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_sf.setItem(l,j, nvalor)
                            self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(0,j).value)))
                        l=l+1
                        
                elif self.actionColombia.isChecked(): #despliega info filtrada por fecha y colombia
                    for i in self.filtered_date_index:
                        for j in range(num_col):
                            valor = self.datasf[i][j]
                            if j== 29:
                                if valor == 'CN Local Colombia':
                                    self.filtered_dr_index.append(i)
                                    self.filtered_dr_index.sort()
                    
                    self.tabla_sf.setRowCount(len(self.filtered_dr_index))
                    l=0
                    for i in self.filtered_dr_index:
                        
                        for j in range(num_col):
                            if i == 0:
                                self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(i,j).value)))
                            else: 
                                if j == 10 or j == 2 or j == 9:
                                    valor10 =self.datasf[i][j]
                                    y = type(valor10) is float
                                    if y == True:
                                        if valor10 == 1:
                                            nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                        else:
                                            seconds10 = (valor10 - 25569) * 86400.0
                                            xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                            xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                            nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                    else:
                                        valor = str(self.datasf[i][j])
                                        nvalor = QtWidgets.QTableWidgetItem(valor)
                                        self.tabla_sf.setItem(l,j, nvalor)
                                else:
                                    
                                    valor = str(self.datasf[i][j])
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_sf.setItem(l,j, nvalor)
                            self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(0,j).value)))
                        l=l+1
                else:  #despliega info filtrada por fecha   
                    
                    self.tabla_sf.setRowCount(len(self.filtered_date_index))
                    l=0
                    for i in self.filtered_date_index:
                        
                        for j in range(num_col):
                            if i == 0:
                                self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(i,j).value)))
                            else: 
                                if j == 10 or j == 2 or j == 9:
                                    valor10 =self.datasf[i][j]
                                    y = type(valor10) is float
                                    if y == True:
                                        if valor10 == 1:
                                            nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                        else:
                                            seconds10 = (valor10 - 25569) * 86400.0
                                            xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                            xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                            nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                    else:
                                        valor = str(self.datasf[i][j])
                                        nvalor = QtWidgets.QTableWidgetItem(valor)
                                        self.tabla_sf.setItem(l,j, nvalor)
                                else:
                                    
                                    valor = str(self.datasf[i][j])
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_sf.setItem(l,j, nvalor)
                            self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(0,j).value)))
                        l=l+1
    #                
            else:
                #no fecha - si region
                if self.actionRegional.isChecked(): #despliega info filtrada por regional
                    for i in self.filtered_date_index:
                        for j in range(num_col):
                            valor = self.datasf[i][j]
                            if j == 25:
                                if valor == 'Wholesale Regional':
                                    self.filtered_region_index.append(i)
                                    self.filtered_region_index.sort()
                    self.tabla_sf.setRowCount(len(self.filtered_region_index))
                    l=0
                    for i in self.filtered_region_index:
                        
                        for j in range(num_col):
                            if i == 0:
                                self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(i,j).value)))
                            else: 
                                if j == 10 or j == 2 or j == 9:
                                    valor10 =self.datasf[i][j]
                                    y = type(valor10) is float
                                    if y == True:
                                        if valor10 == 1:
                                            nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                        else:
                                            seconds10 = (valor10 - 25569) * 86400.0
                                            xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                            xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                            nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                    else:
                                        valor = str(self.datasf[i][j])
                                        nvalor = QtWidgets.QTableWidgetItem(valor)
                                        self.tabla_sf.setItem(l,j, nvalor)
                                else:
                                    
                                    valor = str(self.datasf[i][j])
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_sf.setItem(l,j, nvalor)
                            self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(0,j).value)))
                        l=l+1
                        
                        
                elif self.actionColombia.isChecked(): #despliega info filtrada colombia
                    for i in self.filtered_date_index:
                        for j in range(num_col):
                            valor = self.datasf[i][j]
                            if j== 29:
                                if valor == 'CN Local Colombia':
                                    self.filtered_region_index.append(i)
                                    self.filtered_region_index.sort()
                    self.tabla_sf.setRowCount(len(self.filtered_region_index))
                    l=0
                    for i in self.filtered_region_index:
                        
                        for j in range(num_col):
                            if i == 0:
                                self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(i,j).value)))
                            else: 
                                if j == 10 or j == 2 or j == 9:
                                    valor10 =self.datasf[i][j]
                                    y = type(valor10) is float
                                    if y == True:
                                        if valor10 == 1:
                                            nvalor10 = QtWidgets.QTableWidgetItem('1/01/1900')
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                        else:
                                            seconds10 = (valor10 - 25569) * 86400.0
                                            xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                            xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                            nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                            self.tabla_sf.setItem(l,j, nvalor10)
                                    else:
                                        valor = str(self.datasf[i][j])
                                        nvalor = QtWidgets.QTableWidgetItem(valor)
                                        self.tabla_sf.setItem(l,j, nvalor)
                                else:
                                    
                                    valor = str(self.datasf[i][j])
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_sf.setItem(l,j, nvalor)
                            self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(0,j).value)))
                        l=l+1                    
                else:  #despliega info completa 
                    
                    for i in range(self.hoja_sf.nrows):
                        for j in range(self.hoja_sf.ncols):
                            if i == 0:
                                self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(i,j).value)))
                            else: 
                                if j == 10 or j == 2 or j == 9:
                                    valor10 =self.hoja_sf.cell(i,j).value
                                    y = type(valor10) is float
                                    if y == True:
                                        seconds10 = (valor10 - 25569) * 86400.0
                                        xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                        xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                        nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                        self.tabla_sf.setItem(i-1,j, nvalor10)
                                    else:
                                        valor = str(self.hoja_sf.cell(i,j).value)
                                        nvalor = QtWidgets.QTableWidgetItem(valor)
                                        self.tabla_sf.setItem(i-1,j, nvalor)
                                else:
                                    
                                    valor = str(self.hoja_sf.cell(i,j).value)
                                    nvalor = QtWidgets.QTableWidgetItem(valor)
                                    self.tabla_sf.setItem(i-1,j, nvalor)
        else: 
            for i in range(num_row):
                for j in range(num_col):
                    if i == 0:
                        self.tabla_sf.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_sf.cell(i,j).value)))
                    else: 
                        if j == 10 or j == 2 or j == 9:
                            valor10 =self.hoja_sf.cell(i,j).value
                            y = type(valor10) is float
                            if y == True:
                                seconds10 = (valor10 - 25569) * 86400.0
                                xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                                xtm= datetime.datetime.utcfromtimestamp(seconds10).month
                                nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                                self.tabla_sf.setItem(i-1,j, nvalor10)
                            else:
                                valor = str(self.hoja_sf.cell(i,j).value)
                                nvalor = QtWidgets.QTableWidgetItem(valor)
                                self.tabla_sf.setItem(i-1,j, nvalor)
                        else:
                            
                            valor = str(self.hoja_sf.cell(i,j).value)
                            nvalor = QtWidgets.QTableWidgetItem(valor)
                            self.tabla_sf.setItem(i-1,j, nvalor)
        self.tabla_sf.resizeColumnsToContents()
        self.tabla_busquedasf = np.empty((self.tabla_sf.rowCount(), self.tabla_sf.columnCount()), dtype=('U100'))
        for x in range(self.tabla_sf.rowCount()):
               for y in range(self.tabla_sf.columnCount()):
                   self.tabla_busquedasf[x][y] = self.tabla_sf.item(x,y).text()


    '''
# =============================================================================
# BUSCA LA CELDA SELECCIONADA
# =============================================================================
    '''

    def cellselect(self, row_sf, column_sf):
        self.row_sf = row_sf
        self.column_sf = column_sf
        self.buscar_button.clicked.connect(self.buscar_clicked)
        
    def buscar_clicked(self):        
# =============================================================================
#         BÚSQUEDA DE IDS REPETIDOS
# =============================================================================
        if 'date_select' in globals():
            self.statusbar.showMessage('Periodo seleccionado: %i/%i' %(date_select.month, date_select.year))
        
        idsf = self.tabla_sf.item(self.row_sf, 12)
        self.idsf = idsf.text()
        nordensf = self.tabla_sf.item(self.row_sf,11)
        self.nordensf = nordensf.text()
        ordersf = self.tabla_sf.item(self.row_sf,4)
        self.ordersf = ordersf.text()
        self.tabla_busqueda = np.empty((self.tabla_sf.rowCount(), self.tabla_sf.columnCount()), dtype=('U100'))

        for x in range(self.tabla_sf.rowCount()):
            for y in range(self.tabla_sf.columnCount()):
                self.tabla_busquedasf[x][y] = self.tabla_sf.item(x,y).text()
                self.tabla_sf.item(x,y).setBackground(QtGui.QColor(255,255,255))
        indexsf= [i for i,x in enumerate(self.tabla_busquedasf) for j,y in enumerate(x) if y == self.idsf]
        indexsfo= [i for i,x in enumerate(self.tabla_busquedasf) for j,y in enumerate(x) if y == self.ordersf]
        
        if len(indexsf) == 1:
            nordensf = self.tabla_sf.item(self.row_sf,11)
            self.nordensf = nordensf.text()
            idsf = self.tabla_sf.item(self.row_sf, 12)
            self.idsf = idsf.text()
            operadorsf = self.tabla_sf.item(self.row_sf, 5)
            self.operadorsf = operadorsf.text()
            lenstrsf = len(self.operadorsf)//2
            if lenstrsf >15:
                font = QtGui.QFont()
                font.setPointSize(8)
                if self.operadorsf[lenstrsf] == ' ':
                    z1sf,z2sf= self.operadorsf[:lenstrsf], self.operadorsf[lenstrsf:]
                    self.prtsf = z1sf + '\n' + z2sf
                    self.operador_sf.setText(self.prtsf)
                else:
                    xsf= [msf.start() for msf in re.finditer(' ', self.operadorsf)]
                    takeClosestsf = lambda numsf,collectionsf:min(collectionsf,key=lambda xsf:abs(xsf-numsf))
                    ysf=takeClosestsf(lenstrsf,xsf)
                    z1sf,z2sf= self.operadorsf[:ysf], self.operadorsf[1+ysf:]
                    self.prtsf = z1sf + '\n' + z2sf
                    self.operador_sf.setText(self.prtsf)
            else:
                self.operador_sf.setText(self.operadorsf)
            terminosf = self.tabla_sf.item(self.row_sf, 30)
            self.terminosf = terminosf.text()
            mrcsf = self.tabla_sf.item(self.row_sf, 22)
            self.mrcsf = mrcsf.text()
            nrcsf = self.tabla_sf.item(self.row_sf, 24)
            self.nrcsf = nrcsf.text()
            self.norden_sf.setText(self.nordensf)
            self.id_sf.setText(self.idsf)
            self.termino_sf.setText(self.terminosf)
            self.mrc_sf.setText(self.mrcsf)
            self.nrc_sf.setText(self.nrcsf)
        elif len(indexsf) > 1:
            fecha_max=0
            self.mrcsumsf=0
            nrcsuma=0
            for i in indexsf:
                for j in range(self.tabla_sf.columnCount()):
                    self.tabla_sf.item(i,j).setBackground(QtGui.QColor(255,255,0))
            
            
# =============================================================================
# FECHA MÁS RECIENTE             
# =============================================================================
            for i in indexsf:
                fechastr = self.tabla_busquedasf[i][10]
                fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                if fechastamp > fecha_max:
                    fecha_max= fechastamp
                    fecha_max_index = i
            orden= self.tabla_busquedasf[fecha_max_index][4] 
            index_fecha = [i for i,x in enumerate(self.tabla_busquedasf) for j,y in enumerate(x) if y == orden]
            indexsf= [i for i,x in enumerate(self.tabla_busquedasf) for j,y in enumerate(x) if y == self.idsf]      
            for i in range (len(index_fecha)):
                if self.idsf ==  self.tabla_busquedasf[index_fecha[i]][12]:
                    if self.nordensf == self.tabla_busquedasf[index_fecha[i]][11]:
                        valor = float(self.tabla_busquedasf[index_fecha[i]][22])
                        self.mrcsumsf = valor + self.mrcsumsf  
                        nrc= float(self.tabla_busquedasf[index_fecha[i]][24])
                    
                        if nrc != 0:
                            nrcsuma=nrcsuma+nrc
                            self.nrcsf = self.tabla_busquedasf[index_fecha[i]][24]
                        else:
                            if nrcsuma==0:
                                self.nrcsf = float(0)
            self.mrcsumsf=str(self.mrcsumsf)
            self.nrcsf=str(self.nrcsf)    
            
            nordensf = self.tabla_busquedasf[fecha_max_index][11]
            self.nordensf = str(nordensf)

            idsf = self.tabla_busquedasf[fecha_max_index][12]
            self.idsf = str(idsf)
            operadorsf = self.tabla_busquedasf[fecha_max_index][5]
            self.operadorsf = str(operadorsf)
            lenstrsf = len(self.operadorsf)//2
            if lenstrsf >15:
                font = QtGui.QFont()
                font.setPointSize(8)
                if self.operadorsf[lenstrsf] == ' ':
                    z1sf,z2sf= self.operadorsf[:lenstrsf], self.operadorsf[lenstrsf:]
                    self.prtsf = z1sf + '\n' + z2sf
                    self.operador_sf.setText(self.prtsf)
                else:
                    xsf= [msf.start() for msf in re.finditer(' ', self.operadorsf)]
                    takeClosestsf = lambda numsf,collectionsf:min(collectionsf,key=lambda xsf:abs(xsf-numsf))
                    ysf=takeClosestsf(lenstrsf,xsf)
                    z1sf,z2sf= self.operadorsf[:ysf], self.operadorsf[1+ysf:]
                    self.prtsf = z1sf + '\n' + z2sf
                    self.operador_sf.setText(self.prtsf)
            
            else:
                self.operador_sf.setText(self.operadorsf)
                
            terminosf = self.tabla_busquedasf[fecha_max_index][30]
            self.terminosf = str(terminosf)
            self.norden_sf.setText(self.nordensf)
            self.id_sf.setText(self.idsf)
            self.termino_sf.setText(self.terminosf)
            self.mrc_sf.setText(self.mrcsumsf)
            self.nrc_sf.setText(self.nrcsf)
        
# =============================================================================
#         BÚSQUEDA DE ID EN FACTURACION
# =============================================================================
        
        self.tabla_busqueda = np.empty((self.tabla_fact.rowCount(), self.tabla_fact.columnCount()), dtype=('U100'))

        for x in range(self.tabla_fact.rowCount()):
            for y in range(self.tabla_fact.columnCount()):
                self.tabla_busqueda[x][y] = self.tabla_fact.item(x,y).text()
         
        indexfb= [(i,j) for i,x in enumerate(self.tabla_busqueda) for j,y in enumerate(x) if y == self.idsf]
        indexf1= [i for i,x in enumerate(self.tabla_busqueda) for j,y in enumerate(x) if y == self.idsf]
        indexfo= [i for i,x in enumerate(self.tabla_busqueda) for j,y in enumerate(x) if y == self.nordensf]

        for i in range(self.tabla_fact.rowCount()):
            for j in range(self.tabla_fact.columnCount()):
                self.tabla_fact.item(i,j).setBackground(QtGui.QColor(255,255,255))

        
        if len(indexfb) == 1:
            self.tabla_fact.scrollToItem(self.tabla_fact.selectRow(indexfb[0][0]))
            nordenf = self.tabla_fact.item(indexfb[0][0], 4)
            self.nordenf = nordenf.text()
            mrcf = self.tabla_fact.item(indexfb[0][0], 20)
            self.mrcf = mrcf.text()
            if self.nordenf == self.nordensf:
                self.norden_fact.setText(self.nordenf)
                self.mrc_fact.setText(self.mrcf)
            else:
                self.norden_fact.setText('Las órdenes no coinciden')
                self.mrc_fact.setText('Las órdenes no coinciden')

            if self.nordenf == self.nordensf:
                self.mrc_fact.setText(self.mrcf)
            else:
                self.mrc_fact.setText('Las órdenes no coinciden')
            
            idf = self.tabla_fact.item(indexfb[0][0], 16)
            self.idf = idf.text()

            operadorf = self.tabla_fact.item(indexfb[0][0], 2)
            self.operadorf = operadorf.text()
            
            lenstrf = len(self.operadorf)//2
            if lenstrf >15:
                
                font = QtGui.QFont()
                font.setPointSize(8)
                if self.operadorf[lenstrf] == ' ':
                    z1f,z2f= self.operadorf[:lenstrf], self.operadorf[lenstrf:]
                    self.prtf = z1f + '\n' + z2f
                    self.operador_fact.setText(self.prtf)
                else:
                    xf= [mf.start() for mf in re.finditer(' ', self.operadorf)]
                    takeClosestf = lambda numf,collectionf:min(collectionf,key=lambda xf:abs(xf-numf))
                    yf=takeClosestf(lenstrf,xf)
                    z1f,z2f= self.operadorf[:yf], self.operadorf[1+yf:]
                    self.prtf = z1f + '\n' + z2f
                    self.operador_fact.setText(self.prtf)
            
            else:
                self.operador_fact.setText(self.operadorf)
            
            terminof = self.tabla_fact.item(indexfb[0][0], 5)
            self.terminof = terminof.text()
            self.nrcf ='N/A'
            self.id_fact.setText(self.idf)
            self.termino_fact.setText(self.terminof)
            self.nrc_fact.setText(self.nrcf)
            
        elif len(indexfb) == 0:
            self.norden_fact.setText('Not Found')
            self.id_fact.setText('')
            self.operador_fact.setText('')
            self.termino_fact.setText('')
            self.mrc_fact.setText('')
            self.nrc_fact.setText('')
            
        elif len(indexfb)>1:
            self.tabla_fact.scrollToItem(self.tabla_fact.selectRow(indexfb[0][0]))
            for i in indexf1:
                for j in range(self.tabla_fact.columnCount()):
                    self.tabla_fact.item(i,j).setBackground(QtGui.QColor(255,255,0))

            for i in indexfo:
                nordenf = self.tabla_fact.item(i, 4)
                self.nordenf = nordenf.text()
                if self.nordensf == self.nordenf:
                    if len(indexfo) == 1:
                        mrcsumf = self.tabla_fact.item(i, 20)
                        self.mrcsumf = mrcsumf.text()
                        self.nordenf = nordenf.text()
                    elif len(indexfo) > 1:            
                        self.mrcsumf=0
                        for j in range (len(indexfo)):
                            valor = float(self.tabla_busqueda[indexfo[j]][20])
                            self.mrcsumf = valor + self.mrcsumf
                        self.mrcsumf=str(self.mrcsumf)
                        self.nordenf = nordenf.text()
                    elif len(indexfo) ==0:
                        self.mrcsumf = 'No existe el número de orden'
                        self.nordenf = 'No existe el número de orden'

                    
            
            idf = self.tabla_fact.item(indexfb[0][0], 16)
            self.idf = idf.text()
            
            
            operadorf = self.tabla_fact.item(indexfb[0][0], 2)
            self.operadorf = operadorf.text()
            
            lenstrf = len(self.operadorf)//2
            if lenstrf >15:
                font = QtGui.QFont()
                font.setPointSize(8)
                if self.operadorf[lenstrf] == ' ':
                    z1f,z2f= self.operadorf[:lenstrf], self.operadorf[lenstrf:]
                    self.prtf = z1f + '\n' + z2f
                    self.operador_fact.setText(self.prtf)
                else:
                    xf= [mf.start() for mf in re.finditer(' ', self.operadorf)]
                    takeClosestf = lambda numf,collectionf:min(collectionf,key=lambda xf:abs(xf-numf))
                    yf=takeClosestf(lenstrf,xf)
                    z1f,z2f= self.operadorf[:yf], self.operadorf[1+yf:]
                    self.prtf = z1f + '\n' + z2f
                    self.operador_fact.setText(self.prtf)
            
            else:
                self.operador_fact.setText(self.operadorf)

            
            terminof = self.tabla_fact.item(indexfb[0][0], 5)
            self.terminof = terminof.text()
            self.nrcf ='N/A'
            self.norden_fact.setText(self.nordenf)
            self.id_fact.setText(self.idf)
            self.termino_fact.setText(self.terminof)
            self.mrc_fact.setText(self.mrcsumf)
            self.nrc_fact.setText(self.nrcf)
            
# =============================================================================
#             LABELS BASE
# =============================================================================
        
        indexb= [(i,j) for i,x in enumerate(self.datab) for j,y in enumerate(x) if y == self.idsf]
        indexb1= [i for i,x in enumerate(self.datab) for j,y in enumerate(x) if y == self.idsf]
        indexbo= [i for i,x in enumerate(self.datab) for j,y in enumerate(x) if y == self.nordensf]
        for i in range(self.tabla_base.rowCount()):
            for j in range(self.tabla_base.columnCount()):
                self.tabla_base.item(i,j).setBackground(QtGui.QColor(255,255,255))
        if len(indexb) == 1:
            self.tabla_base.scrollToItem(self.tabla_base.selectRow(indexb[0][0]-3))
            nordenb = self.tabla_base.item(indexb[0][0]-3, 2)
            self.nordenb = nordenb.text()
            idb = self.tabla_base.item(indexb[0][0]-3, 0)
            self.idb = idb.text()
            operadorb = self.tabla_base.item(indexb[0][0]-3, 3)
            self.operadorb = operadorb.text()
            lenstrb = len(self.operadorb)//2
            if lenstrb >15:
                font = QtGui.QFont()
                font.setPointSize(8)
                if self.operadorb[lenstrb] == ' ':
                    z1b,z2b= self.operadorb[:lenstrb], self.operadorb[lenstrb:]
                    self.prtb = z1b + '\n' + z2b
                    self.operador_base.setText(self.prtb)
                else:
                    xb= [mb.start() for mb in re.finditer(' ', self.operadorb)]
                    takeClosestb = lambda numb,collectionb:min(collectionb,key=lambda xb:abs(xb-numb))
                    yb=takeClosestb(lenstrb,xb)
                    z1b,z2b= self.operadorb[:yb], self.operadorb[1+yb:]
                    self.prtb = z1b + '\n' + z2b
                    self.operador_base.setText(self.prtb)
            else:
                self.operador_base.setText(self.operadorb)
            
            terminob = self.tabla_base.item(indexb[0][0]-3, 16)
            self.terminob = terminob.text()
            mrcb = self.tabla_base.item(indexb[0][0]-3, 7)
            self.mrcb = mrcb.text()
            self.nrcb ='N/A'
            self.norden_base.setText(self.nordenb)
            self.id_base.setText(self.idb)
            self.termino_base.setText(self.terminob)
            self.mrc_base.setText(self.mrcb)
            self.nrc_base.setText(self.nrcb)
            
        elif len(indexb) == 0:
            self.norden_base.setText('Not Found')
            self.id_base.setText('')
            self.operador_base.setText('')
            self.termino_base.setText('')
            self.mrc_base.setText('')
            self.nrc_base.setText('')
            
        elif len(indexb)>1:
            for i in indexb1:
                for j in range(self.tabla_base.columnCount()):
                    self.tabla_base.item(i-3,j).setBackground(QtGui.QColor(255,255,0))
            
            self.mrcsumb=0
            nordenb = self.tabla_base.item(i-3, 2)
            self.nordenb = nordenb.text()
            idba = self.tabla_base.item(i-3, 0)
            self.idba = idba.text()
            
            if self.idsf == self.idba:
                print("IDs Iguales")
                print (self.idsf, self.idba)
                print (self.nordensf, self.nordenb)
                if len(indexbo) == 1:
                    mrcsumb = self.tabla_base.item(i-3, 7)
                    self.mrcsumb = mrcsumb.text()
                    self.nordenb = nordenb.text()
                elif len(indexbo) > 1:            
                    self.mrcsumb=0
                    for j in indexbo:
                        idba = self.tabla_base.item(j-3, 0)
                        self.idba = idba.text()
                        print (self.idsf, self.idba)
                        if self.idsf == self.idba:
                            valor = float(self.datab[j][7])
                            print(valor)
                            self.mrcsumb = valor + self.mrcsumb
                            print(self.mrcsumb)
                    self.mrcsumb=str(self.mrcsumb)
                    self.nordenb = nordenb.text()
                elif len(indexbo) ==0:
                    self.mrcsumb = 'No existe el número de orden'
                    self.nordenb = 'No existe el número de orden'

            self.tabla_base.scrollToItem(self.tabla_base.selectRow(indexb[0][0]-3))
            idb = self.tabla_base.item(indexb[0][0]-3, 0)
            self.idb = idb.text()
            operadorb = self.tabla_base.item(indexb[0][0]-3, 3)
            self.operadorb = operadorb.text()
            
            lenstrb = len(self.operadorb)//2
            if lenstrb >15:
                font = QtGui.QFont()
                font.setPointSize(8)
                if self.operadorb[lenstrb] == ' ':
                    z1b,z2b= self.operadorb[:lenstrb], self.operadorb[lenstrb:]
                    self.prtb = z1b + '\n' + z2b
                    self.operador_base.setText(self.prtb)
                else:
                    xb= [mb.start() for mb in re.finditer(' ', self.operadorb)]
                    takeClosestb = lambda numb,collectionb:min(collectionb,key=lambda xb:abs(xb-numb))
                    yb=takeClosestb(lenstrb,xb)
                    z1b,z2b= self.operadorb[:yb], self.operadorb[1+yb:]
                    self.prtb = z1b + '\n' + z2b
                    self.operador_base.setText(self.prtb)
            
            else:
                self.operador_base.setText(self.operadorb)
            
            
            
            terminob = self.tabla_base.item(indexb[0][0]-3, 16)
            self.terminob = terminob.text()
            self.nrcb ='N/A'
            self.norden_base.setText(self.nordenb)
            self.id_base.setText(self.idb)
            self.termino_base.setText(self.terminob)
            self.mrc_base.setText(str(self.mrcsumb))
            self.nrc_base.setText(str(self.nrcb))

        

    """
# =============================================================================
# FILTRO POR TIPO DE SERVICIO
# =============================================================================
    """      
        
    def selectsf(self):
        item = self.select_sf.currentText()
        if item == 'Nuevos Servicios':
                   
           indexsf= [i for i,x in enumerate(self.tabla_busquedasf) for j,y in enumerate(x) if y == 'New Service']
           self.fdatasfr= np.empty((len(indexsf),self.hoja_sf.ncols), dtype=('U100'))
           for k in range(len(indexsf)):
               for l in range(self.hoja_sf.ncols):
                    self.fdatasfr[k][l]=self.tabla_busquedasf[indexsf[k]][l]
           self.tabla_sf.setRowCount(len(indexsf))
           self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
           for i in range(len(indexsf)):
               for j in range(self.hoja_sf.ncols):
                   if j == 10 or j == 2 or j == 9:
                       valor =  self.fdatasfr[i][j]
                       if valor == '':
                           valor =  self.fdatasfr[i][j]
                           self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
                       else:
                           nvalor = QtWidgets.QTableWidgetItem(valor)
                           self.tabla_sf.setItem(i,j, nvalor)
                   else:
                       valor =  self.fdatasfr[i][j]
                       self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
        elif item == 'Novedades':

           indexsf= [i for i,x in enumerate(self.tabla_busquedasf) for j,y in enumerate(x) if y == 'Downgrade' or y == 'Migration' or y=='Reconfiguration' or y=='Upgrade']
           self.fdatasfr= np.empty((len(indexsf),self.hoja_sf.ncols), dtype=('U100'))
           for k in range(len(indexsf)):
               for l in range(self.hoja_sf.ncols):
                    self.fdatasfr[k][l]=self.tabla_busquedasf[indexsf[k]][l]
           self.tabla_sf.setRowCount(len(indexsf))
           self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
           for i in range(len(indexsf)):
               for j in range(self.hoja_sf.ncols):
                   if j == 10 or j == 2 or j == 9:
                       valor =  self.fdatasfr[i][j]
                       if valor == '':
                           valor =  self.fdatasfr[i][j]
                           self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
                       else:
                           nvalor = QtWidgets.QTableWidgetItem(valor)
                           self.tabla_sf.setItem(i,j, nvalor)
                   else:
                       valor =  self.fdatasfr[i][j]
                       self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
        elif item == 'Renovaciones':

                   
           indexsf= [i for i,x in enumerate(self.tabla_busquedasf) for j,y in enumerate(x) if y == 'Renewal']
           self.fdatasfr= np.empty((len(indexsf),self.hoja_sf.ncols), dtype=('U100'))
           for k in range(len(indexsf)):
               for l in range(self.hoja_sf.ncols):
                    self.fdatasfr[k][l]=self.tabla_busquedasf[indexsf[k]][l]
           self.tabla_sf.setRowCount(len(indexsf))
           self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
           for i in range(len(indexsf)):
               for j in range(self.hoja_sf.ncols):
                   if j == 10 or j == 2 or j == 9:
                       valor =  self.fdatasfr[i][j]
                       if valor == '':
                           valor =  self.fdatasfr[i][j]
                           self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
                       else:
                           nvalor = QtWidgets.QTableWidgetItem(valor)
                           self.tabla_sf.setItem(i,j, nvalor)
                   else:
                       valor =  self.fdatasfr[i][j]
                       self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
        elif item == 'Cancelaciones':
                   
            indexsf= [i for i,x in enumerate(self.tabla_busquedasf) for j,y in enumerate(x) if y == 'Disconnection']
            self.fdatasfr= np.empty((len(indexsf),self.hoja_sf.ncols), dtype=('U100'))
            for k in range(len(indexsf)):
                for l in range(self.hoja_sf.ncols):
                        self.fdatasfr[k][l]=self.tabla_busquedasf[indexsf[k]][l]
            self.tabla_sf.setRowCount(len(indexsf))
            self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
            for i in range(len(indexsf)):
                for j in range(self.hoja_sf.ncols):
                    if j == 10 or j == 2 or j == 9:
                        valor =  self.fdatasfr[i][j]
                        if valor == '':
                            valor =  self.fdatasfr[i][j]
                            self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
                        else:
                            nvalor = QtWidgets.QTableWidgetItem(valor)
                            self.tabla_sf.setItem(i,j, nvalor)
                    else:
                        valor =  self.fdatasfr[i][j]
                        self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
        elif item == ' ':
            
            self.tabla_sf.setRowCount(len(self.tabla_busquedasf))
            self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
            for i in range(len(self.tabla_busquedasf)):
                for j in range(self.hoja_sf.ncols):
                    if j == 10 or j == 2 or j == 9:
                        valor =  self.tabla_busquedasf[i][j]
                        if valor == '':
                            valor =  self.tabla_busquedasf[i][j]
                            self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
                        else:
                            nvalor = QtWidgets.QTableWidgetItem(valor)
                            self.tabla_sf.setItem(i,j, nvalor)
                    else:
                        valor =  self.tabla_busquedasf[i][j]
                        self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))


    def exp_base(self):
        filename = QtWidgets.QFileDialog.getSaveFileName(self, 'Save File', '', ".xls(*.xls)")
        wbk = xlwt.Workbook()
        self.sheet = wbk.add_sheet("sheet", cell_overwrite_ok=True)

        for i in range(self.tabla_base.rowCount()):
            for j in range(self.tabla_base.columnCount()):
                if i ==0:
                    if j == 0 or j == 1 or j == 2 or j == 3:
                        style = xlwt.easyxf('pattern: pattern solid, fore_colour violet;'
                                        'font: colour white, bold True;')
                        text =str(self.datab[2][j])
                        self.sheet.write(i, j, text, style=style)
                    elif j == 4 or j == 5 or j == 6 or j == 7 or j == 8:
                        style = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue;' 
                                        'font: colour white, bold True;')
                        #sky_blue
                        text =str(self.datab[2][j])
                        self.sheet.write(i, j, text, style=style)
                    elif j == 9 or j == 10 or j == 11 or j == 12 or j == 13:
                        style = xlwt.easyxf('pattern: pattern solid, fore_colour light_orange;'
                                        'font: colour white, bold True;')
                        text =str(self.datab[2][j])
                        self.sheet.write(i, j, text, style=style)
                    else: 
                        style = xlwt.easyxf('pattern: pattern solid, fore_colour violet;'
                                        'font: colour white, bold True;')
                        text =str(self.datab[2][j])
                        self.sheet.write(i, j, text, style=style)
                        
                else:
                    text = str(self.tabla_base.item(i-1, j).text())
                    self.sheet.write(i, j, text)
        
        wbk.save(filename[0])

    """
# =============================================================================
# Buscar todo
# =============================================================================
    """
    
    def buscar_todo(self):
        nrows = self.hoja_sf.nrows - 7
        ncols = self.hoja_sf.ncols
        index_fact_correctos = list()
        index_fact_incorrectos = list()
        self.sf = np.empty((self.tabla_sf.rowCount(), self.tabla_sf.columnCount()), dtype=('U100'))
        for l in range(self.tabla_sf.rowCount()):
            for m in range(self.tabla_sf.columnCount()):
                self.sf[l][m] = self.tabla_sf.item(l,m).text()
                
        self.facturacion = np.empty((self.tabla_fact.rowCount(), self.tabla_fact.columnCount()), dtype=('U100'))
        for n in range(self.tabla_fact.rowCount()):
            for o in range(self.tabla_fact.columnCount()):
                self.facturacion[n][o] = self.tabla_fact.item(n,o).text()
                
        self.correctos = np.empty((self.tabla_sf.rowCount(), self.tabla_sf.columnCount()), dtype=('U100'))
        self.incorrectos = np.empty((self.tabla_sf.rowCount(), self.tabla_sf.columnCount()), dtype=('U100'))
        global correctos
        global incorrectos
        correctos = []
        correctos_y = []
        incorrectos = []
        incorrectos_y = []



        for x in range (len(self.sf)):
            for y in range (2):
                id_sf = self.sf[x][12]
                norden_sf = self.sf[x][11]
                operador_sf = self.sf[x][5]
                indexsf= [i for i,x1 in enumerate(self.sf) for j,y1 in enumerate(x1) if y1 == id_sf]
                mrc_sf = float(self.sf[x][22])
                norden_igual_list=list()
                norden_dif_list =list()
                mrc_suma_fact=0
                if len(indexsf) == 1: #hay un solo CID
                    indexfb= [i for i,x2 in enumerate(self.facturacion) for j,y2 in enumerate(x2) if y2 == id_sf] #Busca CID en Facturación
                    
                    if len(indexfb) == 1: #un solo CID en Facturación
                        norden_f = self.facturacion[indexfb[0]][4]
                        mrc_f = float(self.facturacion[indexfb[0]][20])
                        if norden_f == norden_sf:                            
                            if mrc_sf == mrc_f:
                                if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                    estado_sf = 'Active'
                                elif self.sf[x][17] == 'Disconnection':
                                    estado_sf = 'Disconnected'
                                if y==1:
                                    fechastr = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                    estado_f = self.facturacion[indexfb[0]][9]
                                    pais_f = self.facturacion[indexfb[0]][22]
                                    bw_f = self.facturacion[indexfb[0]][34]
                                    mrc_dif = abs(mrc_f-mrc_sf)
                                    termino = float(self.sf[x][30])
                                    term_s = termino * 30*24*60*60
                                    fecha_inicio = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                    correctos_y.append(id_sf); correctos_y.append(''); correctos_y.append(norden_sf); correctos_y.append(operador_sf)
                                    correctos_y.append(pais_f); correctos_y.append(estado_f); correctos_y.append(bw_f); correctos_y.append(mrc_f); correctos_y.append('USD')
                                    correctos_y.append(self.sf[x][15]); correctos_y.append(self.sf[x][28]); correctos_y.append(mrc_sf); correctos_y.append(self.sf[x][21]); correctos_y.append(estado_sf)
                                    correctos_y.append(mrc_dif); correctos_y.append(fecha_inicio); correctos_y.append(termino); correctos_y.append(fecha_fin)
                                    correctos_y.append(''); correctos_y.append(''); correctos_y.append('');
                                    correctos.append(correctos_y) 
                                    correctos_y = []
                            else:
                                if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                    estado_sf = 'Active'
                                elif self.sf[x][17] == 'Disconnection':
                                    estado_sf = 'Disconnected'
                                if y == 1:
                                    fechastr = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                    estado_f = self.facturacion[indexfb[0]][9]
                                    pais_f = self.facturacion[indexfb[0]][22]
                                    bw_f = self.facturacion[indexfb[0]][34]
                                    mrc_dif = abs(mrc_f-mrc_sf)
                                    termino = float(self.sf[x][30])
                                    term_s = termino * 30*24*60*60
                                    fecha_inicio = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                    incorrectos_y.append(id_sf); incorrectos_y.append(norden_sf); incorrectos_y.append(operador_sf)
                                    incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_f); incorrectos_y.append('USD')
                                    incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                                    incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                                    incorrectos.append(incorrectos_y)
                                    incorrectos_y = []
                                    
                        else:
                            if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                estado_sf = 'Active'
                            elif self.sf[x][17] == 'Disconnection':
                                estado_sf = 'Disconnected'
                            if y == 1:
                                fechastr = self.sf[x][10]
                                fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                estado_f = self.facturacion[indexfb[0]][9]
                                pais_f = self.facturacion[indexfb[0]][22]
                                bw_f = self.facturacion[indexfb[0]][34]
                                mrc_dif = abs(mrc_f-mrc_sf)
                                termino = float(self.sf[x][30])
                                term_s = termino * 30*24*60*60
                                fecha_inicio = self.sf[x][10]
                                fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                incorrectos_y.append(id_sf); incorrectos_y.append(norden_sf); incorrectos_y.append(operador_sf)
                                incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_f); incorrectos_y.append('USD')
                                incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                                incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                                incorrectos.append(incorrectos_y)
                                incorrectos_y = []
                            
                    elif len(indexfb)>1: #Mas de 1 CID en Facturación
                        for z in indexfb:
                            norden_f =self.facturacion[z][4]
                            if norden_f == norden_sf:
                                norden_igual_list.append(z)
                            else:
                                norden_dif_list.append(z)
                        if len(norden_igual_list) > 1:
                            for z in norden_igual_list:
                                mrc_f=float(self.facturacion[z][20])
                                mrc_suma_fact =mrc_f+mrc_suma_fact
                            if mrc_suma_fact == mrc_sf:
                                if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                    estado_sf = 'Active'
                                elif self.sf[x][17] == 'Disconnection':
                                    estado_sf = 'Disconnected'
                                if y == 1:
                                    fechastr = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                    estado_f = self.facturacion[indexfb[0]][9]
                                    pais_f = self.facturacion[indexfb[0]][22]
                                    bw_f = self.facturacion[indexfb[0]][34]
                                    mrc_dif = abs(mrc_suma_fact-mrc_sf)
                                    termino = float(self.sf[x][30])
                                    term_s = termino * 30*24*60*60
                                    fecha_inicio = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                    correctos_y.append(id_sf); correctos_y.append(''); correctos_y.append(norden_sf); correctos_y.append(operador_sf)
                                    correctos_y.append(pais_f); correctos_y.append(estado_f); correctos_y.append(bw_f); correctos_y.append(mrc_suma_fact); correctos_y.append('USD')
                                    correctos_y.append(self.sf[x][15]); correctos_y.append(self.sf[x][28]); correctos_y.append(mrc_sf); correctos_y.append(self.sf[x][21]); correctos_y.append(estado_sf)
                                    correctos_y.append(mrc_dif); correctos_y.append(fecha_inicio); correctos_y.append(termino); correctos_y.append(fecha_fin)
                                    correctos_y.append(''); correctos_y.append(''); correctos_y.append('');
                                    correctos.append(correctos_y) 
                                    correctos_y = []
                            else:
                                if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                    estado_sf = 'Active'
                                elif self.sf[x][17] == 'Disconnection':
                                    estado_sf = 'Disconnected'
                                if y == 1:
                                    fechastr = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                    estado_f = self.facturacion[indexfb[0]][9]
                                    pais_f = self.facturacion[indexfb[0]][22]
                                    bw_f = self.facturacion[indexfb[0]][34]
                                    mrc_dif = abs(mrc_f-mrc_sf)
                                    termino = float(self.sf[x][30])
                                    term_s = termino * 30*24*60*60
                                    fecha_inicio = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                    incorrectos_y.append(id_sf); incorrectos_y.append(norden_sf); incorrectos_y.append(operador_sf)
                                    incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_suma_fact); incorrectos_y.append('USD')
                                    incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                                    incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                                    incorrectos.append(incorrectos_y)
                                    incorrectos_y = []
                        elif len(norden_igual_list) == 1:
                            mrc_f=float(self.facturacion[norden_igual_list[0]][20])
                            if mrc_f == mrc_sf:
                                if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                    estado_sf = 'Active'
                                elif self.sf[x][17] == 'Disconnection':
                                    estado_sf = 'Disconnected'
                                if y == 1:
                                    fechastr = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                    estado_f = self.facturacion[indexfb[0]][9]
                                    pais_f = self.facturacion[indexfb[0]][22]
                                    bw_f = self.facturacion[indexfb[0]][34]
                                    mrc_dif = abs(mrc_f-mrc_sf)
                                    termino = float(self.sf[x][30])
                                    term_s = termino * 30*24*60*60
                                    fecha_inicio = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                    correctos_y.append(id_sf); correctos_y.append(''); correctos_y.append(norden_sf); correctos_y.append(operador_sf)
                                    correctos_y.append(pais_f); correctos_y.append(estado_f); correctos_y.append(bw_f); correctos_y.append(mrc_f); correctos_y.append('USD')
                                    correctos_y.append(self.sf[x][15]); correctos_y.append(self.sf[x][28]); correctos_y.append(mrc_sf); correctos_y.append(self.sf[x][21]); correctos_y.append(estado_sf)
                                    correctos_y.append(mrc_dif); correctos_y.append(fecha_inicio); correctos_y.append(termino); correctos_y.append(fecha_fin)
                                    correctos_y.append(''); correctos_y.append(''); correctos_y.append('');
                                    correctos.append(correctos_y) 
                                    correctos_y = []
                            else:
                                if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                    estado_sf = 'Active'
                                elif self.sf[x][17] == 'Disconnection':
                                    estado_sf = 'Disconnected'
                                if y == 1:
                                    fechastr = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                    estado_f = self.facturacion[indexfb[0]][9]
                                    pais_f = self.facturacion[indexfb[0]][22]
                                    bw_f = self.facturacion[indexfb[0]][34]
                                    mrc_dif = abs(mrc_f-mrc_sf)
                                    termino = float(self.sf[x][30])
                                    term_s = termino * 30*24*60*60
                                    fecha_inicio = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                    incorrectos_y.append(id_sf); incorrectos_y.append(norden_sf); incorrectos_y.append(operador_sf)
                                    incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_f); incorrectos_y.append('USD')
                                    incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                                    incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                                    incorrectos.append(incorrectos_y)
                                    incorrectos_y = []
                        elif len(norden_igual_list) == 0:
                            if len(norden_dif_list) >= 1:
                                for z  in norden_dif_list:
                                    if y==1:
                                        mrc_f=float(self.facturacion[z][20])
                                        norden_f=self.facturacion[z][4]
                                        fechastr = self.sf[x][10]
                                        fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                        estado_f = self.facturacion[z][9]
                                        pais_f = self.facturacion[z][22]
                                        bw_f = self.facturacion[z][34]
                                        mrc_dif = abs(mrc_f-mrc_sf)
                                        termino = float(self.sf[x][30])
                                        term_s = termino * 30*24*60*60
                                        fecha_inicio = self.sf[x][10]
                                        fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                        fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                        incorrectos_y.append(id_sf); incorrectos_y.append(norden_f); incorrectos_y.append(operador_sf)
                                        incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_f); incorrectos_y.append('USD')
                                        incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                                        incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                                        incorrectos.append(incorrectos_y)
                                        incorrectos_y = []
                            else:
                                if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                    estado_sf = 'Active'
                                elif self.sf[x][17] == 'Disconnection':
                                    estado_sf = 'Disconnected'
                                
                                if y == 1:
                                    mrc_f = int(0)
                                    fechastr = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                    estado_f = 'NA'
                                    pais_f = 'NA'
                                    bw_f = 'NA'
                                    mrc_dif = abs(mrc_f-mrc_sf)
                                    termino = float( self.sf[x][30])
                                    term_s = termino *30*24*60*60
                                    fecha_inicio = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                    incorrectos_y.append(id_sf); incorrectos_y.append(norden_sf); incorrectos_y.append(operador_sf)
                                    incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_f); incorrectos_y.append('USD')
                                    incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                                    incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                                    incorrectos.append(incorrectos_y)
                                    incorrectos_y = [] 
                        
                        if len(norden_dif_list) >= 1:
                                for z  in norden_dif_list:
                                    if y == 1:
                                        mrc_f=float(self.facturacion[z][20])
                                        norden_f=self.facturacion[z][4]
                                        fechastr = self.sf[x][10]
                                        fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                        estado_f = self.facturacion[z][9]
                                        pais_f = self.facturacion[z][22]
                                        bw_f = self.facturacion[z][34]
                                        mrc_dif = abs(mrc_f-mrc_sf)
                                        termino = float(self.sf[x][30])
                                        term_s = termino * 30*24*60*60
                                        fecha_inicio = self.sf[x][10]
                                        fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                        fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                        incorrectos_y.append(id_sf); incorrectos_y.append(norden_f); incorrectos_y.append(operador_sf)
                                        incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_f); incorrectos_y.append('USD')
                                        incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                                        incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                                        incorrectos.append(incorrectos_y)
                                        incorrectos_y = []
                        
                    elif len(indexfb) == 0:
                        if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                            estado_sf = 'Active'
                        elif self.sf[x][17] == 'Disconnection':
                            estado_sf = 'Disconnected'
                        
                        if y == 1:
                            mrc_f = int(0)
                            fechastr = self.sf[x][10]
                            fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                            estado_f = 'NA'
                            pais_f = 'NA'
                            bw_f = 'NA'
                            mrc_dif = abs(mrc_f-mrc_sf)
                            termino = float( self.sf[x][30])
                            term_s = termino *30*24*60*60
                            fecha_inicio = self.sf[x][10]
                            fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                            fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                            incorrectos_y.append(id_sf); incorrectos_y.append(norden_sf); incorrectos_y.append(operador_sf)
                            incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_f); incorrectos_y.append('USD')
                            incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                            incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                            incorrectos.append(incorrectos_y)
                            incorrectos_y = []
                        
                elif len(indexsf) > 1: #Mas de un CID SF
                    fecha_max = 0
                    mrc_suma_sf = 0
                    nrc_suma_sf = 0
                    for z in indexsf:
                        fechastr = str(self.sf[z][10])
                        fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                        if fechastamp > fecha_max:
                            fecha_max= fechastamp
                            fecha_max_index = z
                    orden= self.sf[fecha_max_index][4]
                    index_fecha = [i for i,x3 in enumerate(self.sf) for j,y3 in enumerate(x3) if y3 == orden]
                    fecha_max_str = datetime.datetime.utcfromtimestamp(fecha_max).strftime('%d/%m/%Y')
                    bw_sf=0
                    for z in index_fecha:
                        if id_sf ==  self.sf[z][12]:
                            if norden_sf == self.sf[z][11]:
                                if self.sf[z][10] == fecha_max_str:
                                    valor = float(self.sf[z][22])
                                    if self.sf[x][28] == '':
                                        valorbw = float(0)
                                    else:
                                        valorbw = float(self.sf[x][28])
                                        
                                    if valorbw>bw_sf:
                                        bw_sf=valorbw
                                    mrc_suma_sf = valor + mrc_suma_sf 
                                    nrc_sf= float(self.sf[z][24])
                                    if nrc_sf != 0:
                                        nrc_suma_sf = nrc_suma_sf + nrc_sf
                                        nrc_sf = self.sf[z][24]
                                    else:
                                        if nrc_suma_sf==0:
                                            nrc_sf = float(0)
                    
                    indexfb= [i for i,x2 in enumerate(self.facturacion) for j,y2 in enumerate(x2) if y2 == id_sf] #Busca CID en Facturación
                    if len(indexfb) == 1: #un solo CID en Facturación
                        norden_f = self.facturacion[indexfb[0]][4]
                        mrc_f = float(self.facturacion[indexfb[0]][20])
                        if norden_f == norden_sf:
                            if mrc_suma_sf == mrc_f:
                                if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                    estado_sf = 'Active'
                                elif self.sf[x][17] == 'Disconnection':
                                    estado_sf = 'Disconnected'
                                if y == 1:
                                    fechastr = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                    estado_f = self.facturacion[indexfb[0]][9]
                                    pais_f = self.facturacion[indexfb[0]][22]
                                    bw_f = self.facturacion[indexfb[0]][34]
                                    mrc_dif = abs(mrc_f-mrc_suma_sf)
                                    termino = float(self.sf[x][30])
                                    term_s = termino * 30*24*60*60
                                    fecha_inicio = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                    correctos_y.append(id_sf); correctos_y.append(''); correctos_y.append(norden_sf); correctos_y.append(operador_sf)
                                    correctos_y.append(pais_f); correctos_y.append(estado_f); correctos_y.append(bw_f); correctos_y.append(mrc_f); correctos_y.append('USD')
                                    correctos_y.append(self.sf[x][15]); correctos_y.append(bw_sf); correctos_y.append(mrc_suma_sf); correctos_y.append(self.sf[x][21]); correctos_y.append(estado_sf)
                                    correctos_y.append(mrc_dif); correctos_y.append(fecha_inicio); correctos_y.append(termino); correctos_y.append(fecha_fin)
                                    correctos_y.append(''); correctos_y.append(''); correctos_y.append('');
                                    correctos.append(correctos_y) 
                                    correctos_y = []
                            else:
                                if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                    estado_sf = 'Active'
                                elif self.sf[x][17] == 'Disconnection':
                                    estado_sf = 'Disconnected'
                                if y == 1:
                                    fechastr = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                    estado_f = self.facturacion[indexfb[0]][9]
                                    pais_f = self.facturacion[indexfb[0]][22]
                                    bw_f = self.facturacion[indexfb[0]][34]
                                    mrc_dif = abs(mrc_f-mrc_suma_sf)
                                    termino = float(self.sf[x][30])
                                    term_s = termino * 30*24*60*60
                                    fecha_inicio = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                    incorrectos_y.append(id_sf); incorrectos_y.append(norden_sf); incorrectos_y.append(operador_sf)
                                    incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_f); incorrectos_y.append('USD')
                                    incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_suma_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                                    incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                                    incorrectos.append(incorrectos_y)
                                    incorrectos_y = []
                        else:
                            if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                estado_sf = 'Active'
                            elif self.sf[x][17] == 'Disconnection':
                                estado_sf = 'Disconnected'
                            if y == 1:
                                fechastr = self.sf[x][10]
                                fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                estado_f = self.facturacion[indexfb[0]][9]
                                pais_f = self.facturacion[indexfb[0]][22]
                                bw_f = self.facturacion[indexfb[0]][34]
                                mrc_dif = abs(mrc_f-mrc_sf)
                                termino = float(self.sf[x][30])
                                term_s = termino * 30*24*60*60
                                fecha_inicio = self.sf[x][10]
                                fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                incorrectos_y.append(id_sf); incorrectos_y.append(norden_sf); incorrectos_y.append(operador_sf)
                                incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_f); incorrectos_y.append('USD')
                                incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                                incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                                incorrectos.append(incorrectos_y)
                                incorrectos_y = []
                            
                    elif len(indexfb)>1: #Mas de 1 CID en Facturación
                        for z in indexfb:
                            norden_f =self.facturacion[z][4]
                            if norden_f == norden_sf:
                                norden_igual_list.append(z)
                            else:
                                norden_dif_list.append(z)
                        if len(norden_igual_list) > 1:
                            for z in norden_igual_list:
                                mrc_f=float(self.facturacion[z][20])
                                mrc_suma_fact =mrc_f+mrc_suma_fact
                            if mrc_suma_fact == mrc_suma_sf:
                                if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                    estado_sf = 'Active'
                                elif self.sf[x][17] == 'Disconnection':
                                    estado_sf = 'Disconnected'
                                if y == 1:
                                    fechastr = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                    estado_f = self.facturacion[indexfb[0]][9]
                                    pais_f = self.facturacion[indexfb[0]][22]
                                    bw_f = self.facturacion[indexfb[0]][34]
                                    mrc_dif = abs(mrc_suma_fact-mrc_suma_sf)
                                    termino = float(self.sf[x][30])
                                    term_s = termino * 30*24*60*60
                                    fecha_inicio = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                    correctos_y.append(id_sf); correctos_y.append(''); correctos_y.append(norden_sf); correctos_y.append(operador_sf)
                                    correctos_y.append(pais_f); correctos_y.append(estado_f); correctos_y.append(bw_f); correctos_y.append(mrc_suma_fact); correctos_y.append('USD')
                                    correctos_y.append(self.sf[x][15]); correctos_y.append(self.sf[x][28]); correctos_y.append(mrc_suma_sf); correctos_y.append(self.sf[x][21]); correctos_y.append(estado_sf)
                                    correctos_y.append(mrc_dif); correctos_y.append(fecha_inicio); correctos_y.append(termino); correctos_y.append(fecha_fin)
                                    correctos_y.append(''); correctos_y.append(''); correctos_y.append('');
                                    correctos.append(correctos_y) 
                                    correctos_y = []
                                
                            else:
                                if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                    estado_sf = 'Active'
                                elif self.sf[x][17] == 'Disconnection':
                                    estado_sf = 'Disconnected'
                                if y == 1:
                                    fechastr = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                    estado_f = self.facturacion[indexfb[0]][9]
                                    pais_f = self.facturacion[indexfb[0]][22]
                                    bw_f = self.facturacion[indexfb[0]][34]
                                    mrc_dif = abs(mrc_f-mrc_sf)
                                    termino = float(self.sf[x][30])
                                    term_s = termino * 30*24*60*60
                                    fecha_inicio = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                    incorrectos_y.append(id_sf); incorrectos_y.append(norden_sf); incorrectos_y.append(operador_sf)
                                    incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_suma_fact); incorrectos_y.append('USD')
                                    incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_suma_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                                    incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                                    incorrectos.append(incorrectos_y)
                                    incorrectos_y = []
                        elif len(norden_igual_list) == 1:
                            mrc_f=float(self.facturacion[norden_igual_list[0]][20])
                            if mrc_f == mrc_suma_sf:
                                if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                    estado_sf = 'Active'
                                elif self.sf[x][17] == 'Disconnection':
                                    estado_sf = 'Disconnected'
                                if y == 1:
                                    fechastr = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                    estado_f = self.facturacion[indexfb[0]][9]
                                    pais_f = self.facturacion[indexfb[0]][22]
                                    bw_f = self.facturacion[indexfb[0]][34]
                                    mrc_dif = abs(mrc_f-mrc_suma_sf)
                                    termino = float(self.sf[x][30])
                                    term_s = termino * 30*24*60*60
                                    fecha_inicio = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                    correctos_y.append(id_sf); correctos_y.append(''); correctos_y.append(norden_sf); correctos_y.append(operador_sf)
                                    correctos_y.append(pais_f); correctos_y.append(estado_f); correctos_y.append(bw_f); correctos_y.append(mrc_f); correctos_y.append('USD')
                                    correctos_y.append(self.sf[x][15]); correctos_y.append(self.sf[x][28]); correctos_y.append(mrc_suma_sf); correctos_y.append(self.sf[x][21]); correctos_y.append(estado_sf)
                                    correctos_y.append(mrc_dif); correctos_y.append(fecha_inicio); correctos_y.append(termino); correctos_y.append(fecha_fin)
                                    correctos_y.append(''); correctos_y.append(''); correctos_y.append('');
                                    correctos.append(correctos_y) 
                                    correctos_y = []
                                
                            else:
                                if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                    estado_sf = 'Active'
                                elif self.sf[x][17] == 'Disconnection':
                                    estado_sf = 'Disconnected'
                                if y == 1:
                                    fechastr = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                    estado_f = self.facturacion[indexfb[0]][9]
                                    pais_f = self.facturacion[indexfb[0]][22]
                                    bw_f = self.facturacion[indexfb[0]][34]
                                    mrc_dif = abs(mrc_f-mrc_sf)
                                    termino = float(self.sf[x][30])
                                    term_s = termino * 30*24*60*60
                                    fecha_inicio = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                    incorrectos_y.append(id_sf); incorrectos_y.append(norden_sf); incorrectos_y.append(operador_sf)
                                    incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_f); incorrectos_y.append('USD')
                                    incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_suma_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                                    incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                                    incorrectos.append(incorrectos_y)
                                    incorrectos_y = []
                                
                        elif len(norden_igual_list) == 0:
                            if len(norden_dif_list) >= 1:
                                for z  in norden_dif_list:
                                    if y==1:
                                        if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                            estado_sf = 'Active'
                                        elif self.sf[x][17] == 'Disconnection':
                                            estado_sf = 'Disconnected'
                                        mrc_f=float(self.facturacion[z][20])
                                        norden_f=self.facturacion[z][4]
                                        fechastr = self.sf[x][10]
                                        fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                        estado_f = self.facturacion[z][9]
                                        pais_f = self.facturacion[z][22]
                                        bw_f = self.facturacion[z][34]
                                        mrc_dif = abs(mrc_f-mrc_sf)
                                        termino = float(self.sf[x][30])
                                        term_s = termino * 30*24*60*60
                                        fecha_inicio = self.sf[x][10]
                                        fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                        fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                        incorrectos_y.append(id_sf); incorrectos_y.append(norden_f); incorrectos_y.append(operador_sf)
                                        incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_f); incorrectos_y.append('USD')
                                        incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                                        incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                                        incorrectos.append(incorrectos_y)
                                        incorrectos_y = []
                            else:
                                if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                                    estado_sf = 'Active'
                                elif self.sf[x][17] == 'Disconnection':
                                    estado_sf = 'Disconnected'
                                
                                if y == 1:
                                    mrc_f = int(0)
                                    fechastr = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                    estado_f = 'NA'
                                    pais_f = 'NA'
                                    bw_f = 'NA'
                                    mrc_dif = abs(mrc_f-mrc_sf)
                                    termino = float( self.sf[x][30])
                                    term_s = termino *30*24*60*60
                                    fecha_inicio = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                    incorrectos_y.append(id_sf); incorrectos_y.append(norden_sf); incorrectos_y.append(operador_sf)
                                    incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_f); incorrectos_y.append('USD')
                                    incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                                    incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                                    incorrectos.append(incorrectos_y)
                                    incorrectos_y = [] 
                        
                        
                        
                        if len(norden_dif_list) >= 1:
                            for z  in norden_dif_list:
                                if y == 1:
                                    mrc_f=float(self.facturacion[z][20])
                                    norden_f=self.facturacion[z][4]
                                    fechastr = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                                    estado_f = self.facturacion[z][9]
                                    pais_f = self.facturacion[z][22]
                                    bw_f = self.facturacion[z][34]
                                    mrc_dif = abs(mrc_f-mrc_sf)
                                    termino = float(self.sf[x][30])
                                    term_s = termino * 30*24*60*60
                                    fecha_inicio = self.sf[x][10]
                                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                                    incorrectos_y.append(id_sf); incorrectos_y.append(norden_f); incorrectos_y.append(operador_sf)
                                    incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_f); incorrectos_y.append('USD')
                                    incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                                    incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                                    incorrectos.append(incorrectos_y)
                                    incorrectos_y = []
                    
                    elif len(indexfb) == 0:
                        if self.sf[x][17] == 'New Service' or self.sf[x][17] ==  'Renewal' or self.sf[x][17] == 'Reconfiguration' or self.sf[x][17] == 'Upgrade' or self.sf[x][17] == 'Downgrade' or self.sf[x][17] == 'Migration':
                            estado_sf = 'Active'
                        elif self.sf[x][17] == 'Disconnection':
                            estado_sf = 'Disconnected'
                        
                        if y == 1:
                            mrc_f = int(0)
                            fechastr = self.sf[x][10]
                            fechastamp = int(datetime.datetime.strptime(fechastr, '%d/%m/%Y').timestamp())
                            estado_f = 'NA'
                            pais_f = 'NA'
                            bw_f = 'NA'
                            mrc_dif = abs(mrc_f-mrc_sf)
                            termino = float( self.sf[x][30])
                            term_s = termino *30*24*60*60
                            fecha_inicio = self.sf[x][10]
                            fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                            fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                            incorrectos_y.append(id_sf); incorrectos_y.append(norden_sf); incorrectos_y.append(operador_sf)
                            incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_f); incorrectos_y.append('USD')
                            incorrectos_y.append(self.sf[x][15]); incorrectos_y.append(self.sf[x][28]); incorrectos_y.append(mrc_sf); incorrectos_y.append(self.sf[x][21]); incorrectos_y.append(estado_sf)
                            incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                            incorrectos.append(incorrectos_y)
                            incorrectos_y = []
        
        for x in range (len(self.facturacion)):
            id_fact = self.facturacion[x][16]
            index_fact = [i for i,x2 in enumerate(self.sf) for j,y2 in enumerate(x2) if y2 == id_fact]
            
            if len(index_fact) == 0:
                mrc_sf = int(0)
                estado_sf = 'NA'
                pais_sf = 'NA'
                bw_sf = 'NA'
                currency_sf = 'NA'
                mrc_f = float(self.facturacion[x][20])
                mrc_dif = abs(mrc_f-mrc_sf)
                termino = float(self.facturacion[x][5])
                term_s = termino *30*24*60*60
                fecha_inicio = self.facturacion[x][11]
                if fecha_inicio == "1/01/1900":
                    fecha_fin = "NA"
                else:
                    fechastamp = int(datetime.datetime.strptime(fecha_inicio, '%d/%m/%Y').timestamp()) + term_s
                    fecha_fin = datetime.datetime.utcfromtimestamp(fechastamp).strftime('%d/%m/%Y')
                pais_f = self.facturacion[x][22]
                estado_f = self.facturacion[x][22]
                bw_f = self.facturacion[x][34]
                norden = self.facturacion[x][4]
                operador = self.facturacion[x][2]
                incorrectos_y.append(id_fact); incorrectos_y.append(norden); incorrectos_y.append(operador)
                incorrectos_y.append(pais_f); incorrectos_y.append(estado_f); incorrectos_y.append(bw_f); incorrectos_y.append(mrc_f); incorrectos_y.append('USD')
                incorrectos_y.append(pais_sf); incorrectos_y.append(bw_sf); incorrectos_y.append(mrc_sf); incorrectos_y.append(currency_sf); incorrectos_y.append(estado_sf)
                incorrectos_y.append(mrc_dif); incorrectos_y.append(fecha_inicio); incorrectos_y.append(termino); incorrectos_y.append(fecha_fin)
                incorrectos.append(incorrectos_y)
                incorrectos_y = []
           
        repetidos=[]            
        x = len(correctos)
        for i in range (x):
            id_c = correctos[i][0]
            norden_c = correctos[i][1]
            mrcf_c = correctos[i][6]
            indexfb= [i for i,x2 in enumerate(correctos) for j,y2 in enumerate(x2) if y2 == id_c] #Busca CID en Facturación
            if len(indexfb)>1:
                for j in indexfb:
                    if id_c == correctos[j][0] and norden_c == correctos[j][1] and mrcf_c == correctos[j][6]:
                        if j>i:
                            repetidos.append(j)
                            repetidos.sort()
                    
        from  more_itertools import unique_everseen
        repetidos= list(unique_everseen(repetidos))
        for i in repetidos[::-1]:  
            correctos.pop(i)    


        repetidos_i=[]            
        x = len(incorrectos)
        for i in range (x):
            id_i = incorrectos[i][0]
            norden_i = incorrectos[i][1]
            mrcf_i = incorrectos[i][6]
            mrcsf_i = incorrectos[i][10]
            indexfb= [i for i,x2 in enumerate(incorrectos) for j,y2 in enumerate(x2) if y2 == id_i] #Busca CID en Facturación
            if len(indexfb)>1:
                for j in indexfb:
                    if id_i == incorrectos[j][0] and norden_i == incorrectos[j][1] and mrcf_i == incorrectos[j][6] and mrcsf_i == incorrectos[i][10]:
                        if j>i:
                            repetidos_i.append(j)
                            repetidos_i.sort()
                  
        from  more_itertools import unique_everseen
        repetidos_i= list(unique_everseen(repetidos_i))
        for i in repetidos_i[::-1]:
            incorrectos.pop(i)                        
        
        self.pop3 = PopUp_Incorrectos()
        self.pop3.setObjectName("Incorrectos")
        self.pop3.show()
    
    
    
    def agregar_base(self):
        
        self.base = np.empty((self.tabla_base.rowCount(), self.tabla_base.columnCount()), dtype=('U100'))
        for n in range(self.tabla_base.rowCount()):
            for o in range(self.tabla_base.columnCount()):
                self.base[n][o] = self.tabla_base.item(n,o).text()
        
        global base
        base = self.base.tolist()
        
        for i in range(len(correctos)):
            base.append(correctos[i])

        self.tabla_base.setRowCount(len(base))
        self.tabla_base.setColumnCount(len(base[0]))
        self.header = self.tabla_base.horizontalHeader()
        self.tabla_base.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        for i in range(len(base)):
            for j in range(len(base[0])):
                valor = str(str(base[i][j]))
                nvalor = QtWidgets.QTableWidgetItem(valor)
                self.tabla_base.setItem(i,j, nvalor)
        self.tabla_base.resizeColumnsToContents()
        
        
        
    """
# =============================================================================
# POPUP
# =============================================================================
    """
    
    def openPopUp(self):
        if self.actionSeleccionar_Fecha.isChecked():
          self.pop = MyPopup()
          self.pop.show()
            
    def openWindow1(self):
        filename = QtWidgets.QFileDialog.getSaveFileName(self, 'Guardar Correctos', '', ".xls(*.xls)")
        wbk = xlwt.Workbook()
        self.sheet = wbk.add_sheet("sheet", cell_overwrite_ok=True)
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.bold =True
        style.font = font
        self.sheet.write(0, 0, 'ID', style=style); self.sheet.write(0, 1, 'GP', style=style);self.sheet.write(0, 2, '# Order', style=style); self.sheet.write(0, 3, 'Operador', style=style)
        self.sheet.write(0, 4, 'Pais', style=style); self.sheet.write(0, 5, 'Estado', style=style); self.sheet.write(0, 6, 'BW', style=style)
        self.sheet.write(0, 7, 'MRC', style=style); self.sheet.write(0, 8, 'Currency', style=style); 
        self.sheet.write(0, 9, 'Pais', style=style); self.sheet.write(0, 10, 'BW', style=style); self.sheet.write(0, 11, 'MRC', style=style); self.sheet.write(0, 12, 'Currency', style=style); self.sheet.write(0, 13, 'Estado', style=style)
        self.sheet.write(0, 14, 'Diferencia', style=style); self.sheet.write(0, 15, 'Fecha Inicio', style=style)
        self.sheet.write(0, 16, 'Termino', style=style); self.sheet.write(0, 17, 'Fecha Fin', style=style)
        self.sheet.write(0, 18, '', style=style); self.sheet.write(0, 19, 'ID Tercero', style=style); self.sheet.write(0, 20, 'Valor Tercero', style=style);
        
        l=1
        for i in range(len(base)):
            for j in range(len(base[0])):
                text = str(base[i][j])
                self.sheet.write(l, j, text)
            l=l+1
        
        wbk.save(filename[0])
        
    def openW2(self):
        filename = QtWidgets.QFileDialog.getSaveFileName(self, 'Guardar Incorrectos', '', ".xls(*.xls)")
        wbk = xlwt.Workbook()
        self.sheet = wbk.add_sheet("sheet", cell_overwrite_ok=True)
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.bold =True
        style.font = font
               
        
        self.sheet.write(0, 0, 'ID', style=style); self.sheet.write(0, 1, '# Order', style=style); self.sheet.write(0, 2, 'Operador', style=style)
        self.sheet.write(0, 3, 'Pais', style=style); self.sheet.write(0, 4, 'Estado', style=style); self.sheet.write(0, 5, 'BW', style=style)
        self.sheet.write(0, 6, 'MRC', style=style); self.sheet.write(0, 7, 'Currency', style=style); self.sheet.write(0, 8, 'Pais', style=style)
        self.sheet.write(0, 9, 'BW', style=style); self.sheet.write(0, 10, 'MRC', style=style); self.sheet.write(0, 11, 'Currency', style=style)
        self.sheet.write(0, 12, 'Estado', style=style); self.sheet.write(0, 13, 'Diferencia', style=style); self.sheet.write(0, 14, 'Fecha Inicio', style=style)
        self.sheet.write(0, 15, 'Termino', style=style); self.sheet.write(0, 16, 'Fecha Fin', style=style)
        l=1
        for i in range(len(incorrectos)):
            for j in range(len(incorrectos[0])):
                text = str(incorrectos[i][j])
                self.sheet.write(l, j, text)
            l=l+1
        
        wbk.save(filename[0])
        
    
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label_3.setText(_translate("MainWindow", "Archivo Salesforce"))
        self.exportar_sf.setText(_translate("MainWindow", "Exportar"))
        self.label_29.setText(_translate("MainWindow", "ID"))
        self.select_sf.setItemText(0, _translate("MainWindow", " "))
        self.select_sf.setItemText(1, _translate("MainWindow", "Nuevos Servicios"))
        self.select_sf.setItemText(2, _translate("MainWindow", "Novedades"))
        self.select_sf.setItemText(3, _translate("MainWindow", "Renovaciones"))
        self.select_sf.setItemText(4, _translate("MainWindow", "Cancelaciones"))
        self.label_35.setText(_translate("MainWindow", "Número\n"
"de orden"))
        self.label_33.setText(_translate("MainWindow", "Operador"))
        self.label.setText(_translate("MainWindow", "Archivo Base"))
        self.label_4.setText(_translate("MainWindow", "Número\n"
"de orden"))
        self.exportar_base.setText(_translate("MainWindow", "Exportar"))
        self.label_8.setText(_translate("MainWindow", "MRC"))
        self.label_7.setText(_translate("MainWindow", "Término"))
        self.label_5.setText(_translate("MainWindow", "ID"))
        self.label_9.setText(_translate("MainWindow", "NRC"))
        self.label_6.setText(_translate("MainWindow", "Operador"))
        self.label_36.setText(_translate("MainWindow", "NRC"))
        self.exportar_fact.setText(_translate("MainWindow", "Exportar"))
        self.label_2.setText(_translate("MainWindow", "Archivo Facturación"))
        self.label_17.setText(_translate("MainWindow", "NRC"))
        self.label_21.setText(_translate("MainWindow", "ID"))
        self.label_32.setText(_translate("MainWindow", "MRC"))
        self.label_23.setText(_translate("MainWindow", "MRC"))
        self.label_20.setText(_translate("MainWindow", "Operador"))
        self.label_25.setText(_translate("MainWindow", "Número\n"
"de orden"))
        self.label_19.setText(_translate("MainWindow", "Término"))
        self.label_31.setText(_translate("MainWindow", "Término"))
        self.buscar_button.setText(_translate("MainWindow", "Buscar"))
        self.agregar_button.setText(_translate("MainWindow", "Agregar"))
        self.modificar_button.setText(_translate("MainWindow", "Modificar"))
        self.actualizar_button.setText(_translate("MainWindow", "Actualizar"))
        self.menuMenu.setTitle(_translate("MainWindow", "Menu"))
        self.menuCargar.setTitle(_translate("MainWindow", "Cargar"))
        self.actionCargar_Archivo_Base.setText(_translate("MainWindow", "Cargar Archivo Base"))
        self.actionCargar_Archivo_Facturacion.setText(_translate("MainWindow", "Cargar Archivo Facturación"))
        self.actionCargar_Archivo_Salesforce.setText(_translate("MainWindow", "Cargar Archivo Salesforce"))
        self.actionCambiar_Usuario.setText(_translate("MainWindow", "Cambiar Usuario"))
        self.actionExportar_todo.setText(_translate("MainWindow", "Exportar todo"))
        self.actionSalir.setText(_translate("MainWindow", "Salir"))
        self.actionManual.setText(_translate("MainWindow", "Manual"))
        self.label.setText(_translate("MainWindow", "Archivo Base"))
        self.menuSettings.setTitle(_translate("MainWindow", "Settings"))
        self.actionSeleccionar_Fecha.setText(_translate("MainWindow", "Seleccionar Fecha"))
        self.menuRegion.setTitle(_translate("MainWindow", "Región"))
        self.actionRegional.setText(_translate("MainWindow", "Regional"))
        self.actionColombia.setText(_translate("MainWindow", "Colombia"))


class MyPopup(QtWidgets.QWidget):
    def __init__(self):
        QtWidgets.QWidget.__init__(self)
        import RAP_rc
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/Register/logo2.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.setWindowIcon(icon)
        self.setFixedSize(392, 241)
        self.setWindowTitle('Fecha')
        self.calendarWidget = QtWidgets.QCalendarWidget(self)
        self.calendarWidget.setGeometry(QtCore.QRect(0, 0, 392, 241))
        self.calendarWidget.setObjectName("calendarWidget")
        self.calendarWidget.clicked.connect(self.select_date)
    
        
    def select_date(self):
        global date_select
        date_select=self.calendarWidget.selectedDate().toPyDate()
        ui.statusbar.showMessage('Periodo seleccionado: %i/%i' %(date_select.month, date_select.year))
        self.close()
        return date_select
        
    
class PopUp_Incorrectos(QtWidgets.QWidget):
    def __init__(self):
        QtWidgets.QWidget.__init__(self)
        import RAP_rc
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/Register/logo2.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.setWindowIcon(icon)
        self.setFixedSize(963, 700)
        self.setWindowTitle('Incorrectos')
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        self.widget = QtWidgets.QWidget(self.centralwidget)
        self.widget.setGeometry(QtCore.QRect(0, 0, 961, 681))
        self.widget.setObjectName("widget")
        self.gridLayout = QtWidgets.QGridLayout(self.widget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.scrollArea = QtWidgets.QScrollArea(self.widget)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.scrollArea.setObjectName("scrollArea")
        self.scrollAreaWidgetContents = QtWidgets.QWidget()
        self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 957, 642))
        self.scrollAreaWidgetContents.setObjectName("scrollAreaWidgetContents")
        self.tableWidget = QtWidgets.QTableWidget(self.scrollAreaWidgetContents)
        self.tableWidget.setGeometry(QtCore.QRect(0, 0, 961, 621))
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.scrollArea.setWidget(self.scrollAreaWidgetContents)
        self.gridLayout.addWidget(self.scrollArea, 0, 0, 1, 1)
        self.pushButton = QtWidgets.QPushButton(self.widget)
        self.pushButton.setMinimumSize(QtCore.QSize(93, 28))
        self.pushButton.setMaximumSize(QtCore.QSize(93, 28))
        self.pushButton.setText('Exportar')
        self.gridLayout.addWidget(self.pushButton, 1, 0, 1, 1)
        if len(incorrectos)> 0:
            num_row = len(incorrectos)
            num_col = len(incorrectos[0])
            self.tableWidget.setRowCount(num_row)
            self.tableWidget.setColumnCount(num_col)
            self.header = self.tableWidget.horizontalHeader()
            self.tableWidget.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
            for i in range(num_row):
                for j in range(num_col):
                    header=['ID', '# Order','Operador', 'Pais', 'Estado', 'BW', 'MRC', 'Currency', 'Pais', 'BW', 'MRC', 'Currency', 'Estado', 'Diferencia', 'Fecha Inicio', 'Termino', 'Fecha Fin']
                    self.tableWidget.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(header[j]))
                    valor = str(incorrectos[i][j])
                    nvalor = QtWidgets.QTableWidgetItem(valor)
                    self.tableWidget.setItem(i,j, nvalor)
            self.tableWidget.resizeColumnsToContents()
            
        self.pushButton.clicked.connect(self.save_incorrectos)
        
    def save_incorrectos(self):
        filename = QtWidgets.QFileDialog.getSaveFileName(self, 'Guardar Incorrectos', '', ".xls(*.xls)")
        wbk = xlwt.Workbook()
        self.sheet = wbk.add_sheet("sheet", cell_overwrite_ok=True)
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.bold =True
        style.font = font
        
        
        self.sheet.write(0, 0, 'ID', style=style); self.sheet.write(0, 1, '# Order', style=style); self.sheet.write(0, 2, 'Operador', style=style)
        self.sheet.write(0, 3, 'Pais', style=style); self.sheet.write(0, 4, 'Estado', style=style); self.sheet.write(0, 5, 'BW', style=style)
        self.sheet.write(0, 6, 'MRC', style=style); self.sheet.write(0, 7, 'Currency', style=style); self.sheet.write(0, 8, 'Pais', style=style)
        self.sheet.write(0, 9, 'BW', style=style); self.sheet.write(0, 10, 'MRC', style=style); self.sheet.write(0, 11, 'Currency', style=style)
        self.sheet.write(0, 12, 'Estado', style=style); self.sheet.write(0, 13, 'Diferencia', style=style); self.sheet.write(0, 14, 'Fecha Inicio', style=style)
        self.sheet.write(0, 15, 'Termino', style=style); self.sheet.write(0, 16, 'Fecha Fin', style=style)

        l=1
        for i in range(len(incorrectos)):
            for j in range(len(incorrectos[0])):
                text = str(incorrectos[i][j])
                self.sheet.write(l, j, text)
            l=l+1
        
        wbk.save(filename[0])

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
