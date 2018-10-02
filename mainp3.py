# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'main2.ui'
#
# Created by: PyQt5 UI code generator 5.11.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
import xlrd
import re
import pandas as pd
import numpy as np
import datetime

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
#        self.scrollArea.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
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
        font.setPointSize(8)
        self.operador_sf.setFont(font)
        self.operador_sf.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.operador_sf.setText("")
        self.operador_sf.setAlignment(QtCore.Qt.AlignCenter)
        self.operador_sf.setObjectName("operador_sf")
        self.gridLayout.addWidget(self.operador_sf, 21, 2, 1, 1)
        self.tabla_sf = QtWidgets.QTableWidget(self.scrollAreaWidgetContents)
        self.tabla_sf.setMinimumSize(QtCore.QSize(1031, 280))
        self.tabla_sf.setMaximumSize(QtCore.QSize(16777215, 285))
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
        self.tabla_base.setMaximumSize(QtCore.QSize(16777215, 285))
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
        self.radio_regional = QtWidgets.QRadioButton(self.scrollAreaWidgetContents)
        self.radio_regional.setChecked(False)
        self.radio_regional.setObjectName("radio_regional")
        self.gridLayout.addWidget(self.radio_regional, 23, 3, 1, 1)
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
#        self.select_fact = QtWidgets.QComboBox(self.scrollAreaWidgetContents)
#        self.select_fact.setMinimumSize(QtCore.QSize(151, 22))
#        self.select_fact.setMaximumSize(QtCore.QSize(151, 22))
#        self.select_fact.setObjectName("select_fact")
#        self.select_fact.addItem("")
#        self.select_fact.setItemText(0, "")
#        self.select_fact.addItem("")
#        self.select_fact.addItem("")
#        self.select_fact.addItem("")
#        self.gridLayout.addWidget(self.select_fact, 10, 3, 1, 1)
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
        self.tabla_fact.setMaximumSize(QtCore.QSize(16777215, 285))
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
        self.radio_colombia = QtWidgets.QRadioButton(self.scrollAreaWidgetContents)
        self.radio_colombia.setCheckable(True)
        self.radio_colombia.setChecked(False)
        self.radio_colombia.setAutoExclusive(True)
        self.radio_colombia.setObjectName("radio_colombia")
        self.gridLayout.addWidget(self.radio_colombia, 22, 3, 1, 1)
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
        self.groupBox.setMinimumSize(QtCore.QSize(313, 28))
        self.groupBox.setMaximumSize(QtCore.QSize(300, 28))
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
        self.agregar_button.setGeometry(QtCore.QRect(110, 0, 93, 28))
        self.agregar_button.setMinimumSize(QtCore.QSize(93, 28))
        self.agregar_button.setMaximumSize(QtCore.QSize(93, 28))
        self.agregar_button.setObjectName("agregar_button")
        self.modificar_button = QtWidgets.QPushButton(self.groupBox)
        self.modificar_button.setGeometry(QtCore.QRect(220, 0, 93, 28))
        self.modificar_button.setMinimumSize(QtCore.QSize(93, 28))
        self.modificar_button.setMaximumSize(QtCore.QSize(93, 28))
        self.modificar_button.setObjectName("modificar_button")
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
        self.radio_colombia.raise_()
        self.radio_regional.raise_()
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
#        self.select_fact.raise_()
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
        
        self.menuMes = QtWidgets.QMenu(self.menuSettings)
        self.menuMes.setObjectName("menuMes")
        
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
        
        self.actionEnero = QtWidgets.QAction(MainWindow)
        self.actionEnero.setCheckable(True)
        self.actionEnero.setObjectName("actionEnero")
        self.actionFebrero = QtWidgets.QAction(MainWindow)
        self.actionFebrero.setCheckable(True)
        self.actionFebrero.setObjectName("actionFebrero")
        self.actionAbril = QtWidgets.QAction(MainWindow)
        self.actionAbril.setCheckable(True)
        self.actionAbril.setObjectName("actionAbril")
        self.actionMayo = QtWidgets.QAction(MainWindow)
        self.actionMayo.setCheckable(True)
        self.actionMayo.setObjectName("actionMayo")
        self.actionJunio = QtWidgets.QAction(MainWindow)
        self.actionJunio.setCheckable(True)
        self.actionJunio.setObjectName("actionJunio")
        self.actionJulio = QtWidgets.QAction(MainWindow)
        self.actionJulio.setCheckable(True)
        self.actionJulio.setObjectName("actionJulio")
        self.actionAgosto = QtWidgets.QAction(MainWindow)
        self.actionAgosto.setCheckable(True)
        self.actionAgosto.setObjectName("actionAgosto")
        self.actionSeptiembre = QtWidgets.QAction(MainWindow)
        self.actionSeptiembre.setCheckable(True)
        self.actionSeptiembre.setObjectName("actionSeptiembre")
        self.actionOctubre = QtWidgets.QAction(MainWindow)
        self.actionOctubre.setCheckable(True)
        self.actionOctubre.setObjectName("actionOctubre")
        self.actionNoviembre = QtWidgets.QAction(MainWindow)
        self.actionNoviembre.setCheckable(True)
        self.actionNoviembre.setObjectName("actionNoviembre")
        self.actionDiciembre = QtWidgets.QAction(MainWindow)
        self.actionDiciembre.setCheckable(True)
        self.actionDiciembre.setObjectName("actionDiciembre")
        self.actionRegional = QtWidgets.QAction(MainWindow)
        self.actionRegional.setCheckable(True)
        self.actionRegional.setObjectName("actionRegional")
        self.actionColombia = QtWidgets.QAction(MainWindow)
        self.actionColombia.setCheckable(True)
        self.actionColombia.setObjectName("actionColombia")
        self.actionMarzo = QtWidgets.QAction(MainWindow)
        self.actionMarzo.setCheckable(True)
        self.actionMarzo.setObjectName("actionMarzo")
        
        
        
        self.menuMenu.addAction(self.actionCambiar_Usuario)
        self.menuMenu.addAction(self.actionExportar_todo)
        self.menuMenu.addAction(self.actionManual)
        self.menuMenu.addAction(self.actionSalir)
        
        self.menuCargar.addAction(self.actionCargar_Archivo_Base)
        self.menuCargar.addAction(self.actionCargar_Archivo_Facturacion)
        self.menuCargar.addAction(self.actionCargar_Archivo_Salesforce)
        self.menubar.addAction(self.menuMenu.menuAction())
        self.menubar.addAction(self.menuCargar.menuAction())
        
        self.menuMes.addAction(self.actionEnero)
        self.menuMes.addAction(self.actionFebrero)
        self.menuMes.addAction(self.actionMarzo)
        self.menuMes.addAction(self.actionAbril)
        self.menuMes.addAction(self.actionMayo)
        self.menuMes.addAction(self.actionJunio)
        self.menuMes.addAction(self.actionJulio)
        self.menuMes.addAction(self.actionAgosto)
        self.menuMes.addAction(self.actionSeptiembre)
        self.menuMes.addAction(self.actionOctubre)
        self.menuMes.addAction(self.actionNoviembre)
        self.menuMes.addAction(self.actionDiciembre)
        self.menuRegion.addAction(self.actionRegional)
        self.menuRegion.addAction(self.actionColombia)
        self.menuSettings.addAction(self.menuMes.menuAction())
        self.menuSettings.addAction(self.menuRegion.menuAction())
        self.menubar.addAction(self.menuSettings.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        
        """
# =============================================================================
#         END
# =============================================================================
        """
        
        
        import RAP_rc
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/Register/logo2.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
#        MainWindow.setWindowState(QtCore.Qt.WindowMaximized)
        MainWindow.setWindowState(MainWindow.windowState() & ~QtCore.Qt.WindowMinimized | QtCore.Qt.WindowActive)
        MainWindow.setFocus(QtCore.Qt.PopupFocusReason)
#        MainWindow.activateWindow()
        MainWindow.raise_()
        self.statusbar.showMessage('User: Test')
        self.actionCargar_Archivo_Base.triggered.connect(self.getxlsbase)
        self.actionCargar_Archivo_Facturacion.triggered.connect(self.getxlsfacturacion)
        self.actionCargar_Archivo_Salesforce.triggered.connect(self.getxlssf)
        self.tabla_sf.cellClicked.connect(self.cellselect)
#        self.select_fact.currentIndexChanged.connect(self.selectfact)
        self.select_sf.currentIndexChanged.connect(self.selectsf)
        self.radio_colombia.toggled.connect(self.radioc)
        self.radio_regional.toggled.connect(self.radior)
#        self.buscar_button.clicked.connect(self.buscar_clicked)
#        self.loop()
        
        
#    def loop(self):
#        self.tabla_sf.cellClicked.connect(self.cellselect)
#
##        cellselect(self, item_base)
##        self.buscar_button.clicked.connect(self.buscar_clicked)  
#        self.select_fact.currentIndexChanged.connect(self.selectfact)
#        self.select_sf.currentIndexChanged.connect(self.selectsf)
#        self.radio_colombia.toggled.connect(self.radioc)
#        self.radio_regional.toggled.connect(self.radior)
#        if self.radio_colombia.isChecked():
#            self.norden_fact.setText('Prueba Radio success')
#            self.id_fact.setText('')
#        else:
#            self.id_fact.setText('Prueba radio failed')
#            self.norden_fact.setText('')

    def radioc(self):
        indexsf= [i for i,x in enumerate(self.datasf) for j,y in enumerate(x) if y == 'Wholesale Local']
        self.fdatasfc= np.empty((len(indexsf),self.hoja_sf.ncols), dtype=('U100'))

        for k in range(len(indexsf)):
            for l in range(self.hoja_sf.ncols):
                 self.fdatasfc[k][l]=self.datasf[indexsf[k]][l]
        
        self.tabla_sf.setRowCount(len(indexsf))
        self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
        for i in range(len(indexsf)):
            for j in range(self.hoja_sf.ncols):
                if j == 10 or j == 2 or j == 9:
                    valor =  self.fdatasfc[i][j]
                    if valor == '':
                        valor =  self.fdatasfc[i][j]
                        self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
                    else:
                        valor = float(valor)
                        y = type(valor) is float
#                            print(y)
                        if y == True:
                            seconds = (valor - 25569) * 86400.0
                            xt= datetime.datetime.utcfromtimestamp(seconds).strftime('%d/%m/%Y')
                            nvalor = QtWidgets.QTableWidgetItem(xt)
                            self.tabla_sf.setItem(i,j, nvalor)
                else:
                    valor =  self.fdatasfc[i][j]
                    self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
#        self.loop()
        

    def radior(self):

        indexsf= [i for i,x in enumerate(self.datasf) for j,y in enumerate(x) if y == 'Wholesale Regional']
        self.fdatasfr= np.empty((len(indexsf),self.hoja_sf.ncols), dtype=('U100'))

        for k in range(len(indexsf)):
            for l in range(self.hoja_sf.ncols):
                 self.fdatasfr[k][l]=self.datasf[indexsf[k]][l]
        
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
                        valor = float(valor)
                        y = type(valor) is float
#                           print(y)
                        if y == True:
                            seconds = (valor - 25569) * 86400.0
                            xt= datetime.datetime.utcfromtimestamp(seconds).strftime('%d/%m/%Y')
                            nvalor = QtWidgets.QTableWidgetItem(xt)
                            self.tabla_sf.setItem(i,j, nvalor)
                else:
                    valor =  self.fdatasfr[i][j]
                    self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
        
#        self.loop()

            
#    def selectfact(self):
##        
#        if self.radio_colombia.isChecked():
#            item = self.select_fact.currentText()
#            if item == 'Nuevos Servicios':
#                indexsf= [i for i,x in enumerate(self.datasf) for j,y in enumerate(x) if y == 'Wholesale Local']
#                self.fdatasfc= np.empty((len(indexsf),self.hoja_sf.ncols), dtype=('U100'))
#
#                for k in range(len(indexsf)):
#                    for l in range(self.hoja_sf.ncols):
#                         self.fdatasfc[k][l]=self.datasf[indexsf[k]][l]
#                
#                self.tabla_sf.setRowCount(len(indexsf))
#                self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
#                for i in range(len(indexsf)):
#                    for j in range(self.hoja_sf.ncols):
#                        valor =  self.fdatasfc[i][j]
#                        self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
#            elif item == 'Novedades':
#                self.nrc_fact.setText('Novedades')
#            elif item == 'Renovaciones':
#                self.nrc_fact.setText('Renovaciones')
#
#        elif self.radio_regional.isChecked():
#            item = self.select_fact.currentText()
#            if item == 'Nuevos Servicios':
#                self.nrc_fact.setText('Nuevos servicios')
#            elif item == 'Novedades':
#                self.nrc_fact.setText('Novedades')
#            elif item == 'Renovaciones':
#                self.nrc_fact.setText('Renovaciones')
        
#        self.loop()
        
    def selectsf(self):
        
        
        item = self.select_sf.currentText()
        if self.radio_colombia.isChecked():
            if item == 'Nuevos Servicios':
                indexsf_ns= [i for i,x in enumerate(self.fdatasfc) for j,y in enumerate(x) if y == 'New Service']
                self.ns_col= np.empty((len(indexsf_ns),self.hoja_sf.ncols), dtype=('U100'))

                for k in range(len(indexsf_ns)):
                    for l in range(self.hoja_sf.ncols):
                         self.ns_col[k][l]=self.fdatasfc[indexsf_ns[k]][l]
                
                self.tabla_sf.setRowCount(len(indexsf_ns))
                self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
                for i in range(len(indexsf_ns)):
                    for j in range(self.hoja_sf.ncols):
                        valor =  self.ns_col[i][j]
                        self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
            elif item == 'Novedades':
                indexsf_n= [i for i,x in enumerate(self.fdatasfc) for j,y in enumerate(x) if y == 'Downgrade' or y == 'Migration' or y=='Reconfiguration' or y=='Upgrade']
                self.n_col= np.empty((len(indexsf_n),self.hoja_sf.ncols), dtype=('U100'))

                for k in range(len(indexsf_n)):
                    for l in range(self.hoja_sf.ncols):
                         self.n_col[k][l]=self.fdatasfc[indexsf_n[k]][l]
                
                self.tabla_sf.setRowCount(len(indexsf_n))
                self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
                for i in range(len(indexsf_n)):
                    for j in range(self.hoja_sf.ncols):
                        valor =  self.n_col[i][j]
                        self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
            elif item == 'Renovaciones':
                indexsf_r= [i for i,x in enumerate(self.fdatasfc) for j,y in enumerate(x) if y == 'Renewal']
                self.r_col= np.empty((len(indexsf_r),self.hoja_sf.ncols), dtype=('U100'))

                for k in range(len(indexsf_r)):
                    for l in range(self.hoja_sf.ncols):
                         self.r_col[k][l]=self.fdatasfc[indexsf_r[k]][l]
                
                self.tabla_sf.setRowCount(len(indexsf_r))
                self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
                for i in range(len(indexsf_r)):
                    for j in range(self.hoja_sf.ncols):
                        valor =  self.r_col[i][j]
                        self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
            elif item == 'Cancelaciones':
                indexsf_c= [i for i,x in enumerate(self.fdatasfc) for j,y in enumerate(x) if y == 'Disconnection']
                self.c_col= np.empty((len(indexsf_c),self.hoja_sf.ncols), dtype=('U100'))

                for k in range(len(indexsf_c)):
                    for l in range(self.hoja_sf.ncols):
                         self.c_col[k][l]=self.fdatasfc[indexsf_c[k]][l]
                
                self.tabla_sf.setRowCount(len(indexsf_c))
                self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
                for i in range(len(indexsf_c)):
                    for j in range(self.hoja_sf.ncols):
                        valor =  self.c_col[i][j]
                        self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
        
            elif item == ' ':
                
#                indexsf_c= [i for i,x in enumerate(self.fdatasfc) for j,y in enumerate(x) if y == 'Disconnection']
                self.c_col= np.empty((len(self.fdatasfc),self.hoja_sf.ncols), dtype=('U100'))

                for k in range(len(self.fdatasfc)):
                    for l in range(self.hoja_sf.ncols):
                         self.c_col[k][l]=self.fdatasfc[k][l]
                
                self.tabla_sf.setRowCount(len(self.fdatasfc))
                self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
                for i in range(len(self.fdatasfc)):
                    for j in range(self.hoja_sf.ncols):
                        valor =  self.c_col[i][j]
                        self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
                
        elif self.radio_regional.isChecked():
            if item == 'Nuevos Servicios':
                indexsf_ns= [i for i,x in enumerate(self.fdatasfr) for j,y in enumerate(x) if y == 'New Service']
                self.ns_r= np.empty((len(indexsf_ns),self.hoja_sf.ncols), dtype=('U100'))

                for k in range(len(indexsf_ns)):
                    for l in range(self.hoja_sf.ncols):
                         self.ns_r[k][l]=self.fdatasfr[indexsf_ns[k]][l]
                
                self.tabla_sf.setRowCount(len(indexsf_ns))
                self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
                for i in range(len(indexsf_ns)):
                    for j in range(self.hoja_sf.ncols):
                        valor =  self.ns_r[i][j]
                        self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
            elif item == 'Novedades':
                indexsf_n= [i for i,x in enumerate(self.fdatasfr) for j,y in enumerate(x) if y == 'Downgrade' or y == 'Migration' or y=='Reconfiguration' or y=='Upgrade']
                self.n_r= np.empty((len(indexsf_n),self.hoja_sf.ncols), dtype=('U100'))

                for k in range(len(indexsf_n)):
                    for l in range(self.hoja_sf.ncols):
                         self.n_r[k][l]=self.fdatasfr[indexsf_n[k]][l]
                
                self.tabla_sf.setRowCount(len(indexsf_n))
                self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
                for i in range(len(indexsf_n)):
                    for j in range(self.hoja_sf.ncols):
                        valor =  self.n_r[i][j]
                        self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
            elif item == 'Renovaciones':
                indexsf_r= [i for i,x in enumerate(self.fdatasfr) for j,y in enumerate(x) if y == 'Renewal']
                self.r_r= np.empty((len(indexsf_r),self.hoja_sf.ncols), dtype=('U100'))

                for k in range(len(indexsf_r)):
                    for l in range(self.hoja_sf.ncols):
                         self.r_r[k][l]=self.fdatasfr[indexsf_r[k]][l]
                
                self.tabla_sf.setRowCount(len(indexsf_r))
                self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
                for i in range(len(indexsf_r)):
                    for j in range(self.hoja_sf.ncols):
                        valor =  self.r_r[i][j]
                        self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
            elif item == 'Cancelaciones':
                indexsf_c= [i for i,x in enumerate(self.fdatasfr) for j,y in enumerate(x) if y == 'Disconnection']
                self.c_r= np.empty((len(indexsf_c),self.hoja_sf.ncols), dtype=('U100'))

                for k in range(len(indexsf_c)):
                    for l in range(self.hoja_sf.ncols):
                         self.c_r[k][l]=self.fdatasfr[indexsf_c[k]][l]
                
                self.tabla_sf.setRowCount(len(indexsf_c))
                self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
                for i in range(len(indexsf_c)):
                    for j in range(self.hoja_sf.ncols):
                        valor =  self.c_r[i][j]
                        self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
             
                
            elif item == ' ':
                
#                indexsf_c= [i for i,x in enumerate(self.fdatasfc) for j,y in enumerate(x) if y == 'Disconnection']
                self.c_col= np.empty((len(self.fdatasfr),self.hoja_sf.ncols), dtype=('U100'))

                for k in range(len(self.fdatasfr)):
                    for l in range(self.hoja_sf.ncols):
                         self.c_col[k][l]=self.fdatasfr[k][l]
                
                self.tabla_sf.setRowCount(len(self.fdatasfr))
                self.tabla_sf.setColumnCount(self.hoja_sf.ncols)
                for i in range(len(self.fdatasfr)):
                    for j in range(self.hoja_sf.ncols):
                        valor =  self.c_col[i][j]
                        self.tabla_sf.setItem(i,j, QtWidgets.QTableWidgetItem(valor))
#        self.loop()
        
    def buscar_clicked(self):
        
# =============================================================================
#         LABELS DE SALESFORCE
# =============================================================================
#        nordensf = self.tabla_sf.item(self.row_sf,11)
#        self.nordensf = nordensf.text()
#        idsf = self.tabla_sf.item(self.row_sf, 12)
#        self.idsf = idsf.text()
#        
#        
#        operadorsf = self.tabla_sf.item(self.row_sf, 5)
#        self.operadorsf = operadorsf.text()
#        lenstrsf = len(self.operadorsf)//2
#        if lenstrsf >15:
#            font = QtGui.QFont()
#            font.setPointSize(8)
#            if self.operadorsf[lenstrsf] == ' ':
#                z1sf,z2sf= self.operadorsf[:lenstrsf], self.operadorsf[lenstrsf:]
#                self.prtsf = z1sf + '\n' + z2sf
#                self.operador_sf.setText(self.prtsf)
#            else:
#                xsf= [msf.start() for msf in re.finditer(' ', self.operadorsf)]
#                takeClosestsf = lambda numsf,collectionsf:min(collectionsf,key=lambda xsf:abs(xsf-numsf))
#                ysf=takeClosestsf(lenstrsf,xsf)
#                z1sf,z2sf= self.operadorsf[:ysf], self.operadorsf[1+ysf:]
#                self.prtsf = z1sf + '\n' + z2sf
#                self.operador_sf.setText(self.prtsf)
#        
#        else:
#            self.operador_sf.setText(self.operadorsf)
#    
#        
#        
#        terminosf = self.tabla_sf.item(self.row_sf, 30)
#        self.terminosf = terminosf.text()
#        mrcsf = self.tabla_sf.item(self.row_sf, 22)
#        self.mrcsf = mrcsf.text()
#        nrcsf = self.tabla_sf.item(self.row_sf, 24)
#        self.nrcsf = nrcsf.text()
#        self.norden_sf.setText(self.nordensf)
#        self.id_sf.setText(self.idsf)
#        
#        self.termino_sf.setText(self.terminosf)
#        self.mrc_sf.setText(self.mrcsf)
#        self.nrc_sf.setText(self.nrcsf)
        
        
# =============================================================================
#         BÚSQUEDA DE IDS REPETIDOS
# =============================================================================
        idsf = self.tabla_sf.item(self.row_sf, 12)
        self.idsf = idsf.text()
        indexsf= [i for i,x in enumerate(self.datasf) for j,y in enumerate(x) if y == self.idsf]
        
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
            
        elif len(indexsf)== 0:
            print('Error')
            
        elif len(indexsf) > 1:
            self.mrcsumsf=0
            for i in range (len(indexsf)):
                self.mrcsumsf = self.datasf[indexsf[i]][22] + self.mrcsumsf
            self.mrcsumsf=str(self.mrcsumsf)
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
#            mrcsf = self.tabla_sf.item(self.row_sf, 22)
#            self.mrcsf = mrcsf.text()
            nrcsf = self.tabla_sf.item(self.row_sf, 24)
            self.nrcsf = nrcsf.text()
            self.norden_sf.setText(self.nordensf)
            self.id_sf.setText(self.idsf)
            
            self.termino_sf.setText(self.terminosf)
            self.mrc_sf.setText(self.mrcsumsf)
            self.nrc_sf.setText(self.nrcsf)
        
# =============================================================================
#         BÚSQUEDA DE ID EN FACTURACION
# =============================================================================
        
        indexf= [(i,j) for i,x in enumerate(self.dataf) for j,y in enumerate(x) if y == self.idsf]
        indexf1= [i for i,x in enumerate(self.dataf) for j,y in enumerate(x) if y == self.idsf]
#        print(indexf)
        if len(indexf) == 1:
            self.tabla_fact.scrollToItem(self.tabla_fact.selectRow(indexf[0][0]-1))
            nordenf = self.tabla_fact.item(indexf[0][0]-1, 4)
            self.nordenf = nordenf.text()
            idf = self.tabla_fact.item(indexf[0][0]-1, 16)
            self.idf = idf.text()
            
            
            operadorf = self.tabla_fact.item(indexf[0][0]-1, 2)
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
#                    print(self.prtf)
                    self.operador_fact.setText(self.prtf)
            
            else:
                self.operador_fact.setText(self.operadorf)
            
            
            
            terminof = self.tabla_fact.item(indexf[0][0]-1, 5)
            self.terminof = terminof.text()
            mrcf = self.tabla_fact.item(indexf[0][0]-1, 20)
            self.mrcf = mrcf.text()
#            nrcb = self.tabla_fact.item(indexf[0][0]-1, 20)
            self.nrcf ='N/A'
#            self.nrcb = nrcb.text()
            self.norden_fact.setText(self.nordenf)
            self.id_fact.setText(self.idf)
            self.termino_fact.setText(self.terminof)
            self.mrc_fact.setText(self.mrcf)
            self.nrc_fact.setText(self.nrcf)
            
        elif len(indexf) == 0:
            
#            self.tabla_fact.scrollToItem(self.tabla_fact.selectRow(indexf[0][0]-1))
            self.norden_fact.setText('Not Found')
            self.id_fact.setText('')
            self.operador_fact.setText('')
            self.termino_fact.setText('')
            self.mrc_fact.setText('')
            self.nrc_fact.setText('')
            
        elif len(indexf)>1:
            self.mrcsumf=0
            for i in range (len(indexf1)):
                self.mrcsumf = self.dataf[indexf1[i]][20] + self.mrcsumf
            self.mrcsumf=str(self.mrcsumf)
            self.tabla_fact.scrollToItem(self.tabla_fact.selectRow(indexf[0][0]-1))
            nordenf = self.tabla_fact.item(indexf[0][0]-1, 4)
            self.nordenf = nordenf.text()
            idf = self.tabla_fact.item(indexf[0][0]-1, 16)
            self.idf = idf.text()
            
            
            operadorf = self.tabla_fact.item(indexf[0][0]-1, 2)
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
#                    print(self.prtf)
                    self.operador_fact.setText(self.prtf)
            
            else:
                self.operador_fact.setText(self.operadorf)

            
            terminof = self.tabla_fact.item(indexf[0][0]-1, 5)
            self.terminof = terminof.text()
#            mrcf = self.tabla_fact.item(indexf[0][0]-1, 20)
#            self.mrcf = mrcf.text()
#            nrcb = self.tabla_fact.item(indexf[0][0]-1, 20)
            self.nrcf ='N/A'
#            self.nrcb = nrcb.text()
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
        
#        print(len(indexb))
#        print(indexb)
        if len(indexb) == 1:
            self.tabla_base.scrollToItem(self.tabla_base.selectRow(indexb[0][0]-3))
            nordenb = self.tabla_base.item(indexb[0][0]-3, 3)
            self.nordenb = nordenb.text()
            idb = self.tabla_base.item(indexb[0][0]-3, 1)
            self.idb = idb.text()
            
            
            operadorb = self.tabla_base.item(indexb[0][0]-3, 4)
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
#                    print(self.prtf)
                    self.operador_base.setText(self.prtb)
            
            else:
                self.operador_base.setText(self.operadorb)
            
            
            
            terminob = self.tabla_base.item(indexb[0][0]-3, 23)
            self.terminob = terminob.text()
            mrcb = self.tabla_base.item(indexb[0][0]-3, 8)
#            print(indexb[0][0]-1, 8)
            self.mrcb = mrcb.text()
#            nrcb = self.tabla_fact.item(indexf[0][0]-1, 20)
            self.nrcb ='N/A'
#            self.nrcb = nrcb.text()
            self.norden_base.setText(self.nordenb)
            self.id_base.setText(self.idb)
            self.termino_base.setText(self.terminob)
            self.mrc_base.setText(self.mrcb)
            self.nrc_base.setText(self.nrcb)
            
        elif len(indexb) == 0:
#            self.tabla_base.scrollToItem(self.tabla_base.selectRow(indexb[0][0]-1))
            self.norden_base.setText('Not Found')
            self.id_base.setText('')
            self.operador_base.setText('')
            self.termino_base.setText('')
            self.mrc_base.setText('')
            self.nrc_base.setText('')
            
        elif len(indexb)>1:
            self.mrcsumb=0
            for i in range (len(indexb1)):
                self.mrcsumb = self.datab[indexb1[i]][8] + self.mrcsumb
            self.mrcsumb=str(self.mrcsumb)
            
#            for i in range (len(indexb1)):
#                for j in range (self.hoja_sf.ncols):
#                    if indexb1[i]>0:
#                        self.tabla_base.item(indexb[0][0]-3,j).setBackground(QtGui.QColor(255,255,0))

            self.tabla_base.scrollToItem(self.tabla_base.selectRow(indexb[0][0]-1))
            nordenb = self.tabla_base.item(indexb[0][0]-3, 3)
            self.nordenb = nordenb.text()
            idb = self.tabla_base.item(indexb[0][0]-3, 1)
            self.idb = idb.text()
            
            
            operadorb = self.tabla_base.item(indexb[0][0]-3, 4)
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
#                    print(self.prtf)
                    self.operador_base.setText(self.prtb)
            
            else:
                self.operador_base.setText(self.operadorb)
            
            
            
            terminob = self.tabla_base.item(indexb[0][0]-3, 23)
            self.terminob = terminob.text()
#            mrcb = self.tabla_base.item(indexb[0][0]-3, 8)
#            self.mrcb = mrcb.text()
            
#            nrcb = self.tabla_fact.item(indexf[0][0]-1, 20)
            self.nrcb ='N/A'
#            self.nrcb = nrcb.text()
            self.norden_base.setText(self.nordenb)
            self.id_base.setText(self.idb)
            self.termino_base.setText(self.terminob)
            self.mrc_base.setText(self.mrcsumb)
            self.nrc_base.setText(self.nrcb)
            
#        for k in range(indexb):
            
        
        
# =============================================================================
#         LABELS ARCHIVO BASE
# =============================================================================
#        nordenb = self.tabla_base.item(self.row_base, 0)
#        self.nordenb = nordenb.text()
#        idb = self.tabla_base.item(self.row_base, 1)
#        self.idb = idb.text()
#        operadorb = self.tabla_base.item(self.row_base, 2)
#        self.operadorb = operadorb.text()
#        terminob = self.tabla_base.item(self.row_base, 3)
#        self.terminob = terminob.text()
#        mrcb = self.tabla_base.item(self.row_base, 4)
#        self.mrcb = mrcb.text()
#        nrcb = self.tabla_base.item(self.row_base, 5)
#        self.nrcb = nrcb.text()
#        self.norden_base.setText(self.nordenb)
#        self.id_base.setText(self.idb)
#        self.operador_base.setText(self.operadorb)
#        self.termino_base.setText(self.terminob)
#        self.mrc_base.setText(self.mrcb)
#        self.nrc_base.setText(self.nrcb)
#        self.loop()
        
    def cellselect(self, row_sf, column_sf):
#        global row_base, column_base
        self.row_sf = row_sf
        self.column_sf = column_sf
#        item_base = self.tabla_base.itemAt(row_base,column_base)
#        num_col = hoja_base.nco
#        return row_base, column_base
#        print(self.tabla_base.item(self.row_base, self.column_base))
#        print('Row %d and Column %d was clicked AAAAA' %(self.row_base, self.column_base))
        self.buscar_button.clicked.connect(self.buscar_clicked)
        
    def getxlsbase(self):
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
                    if j == 17 or j == 19:
                        valor10 =float(self.hoja_base.cell(i,j).value)
                        y = type(valor10) is float
                        print(type(valor10))
                        if y == True:
                            seconds10 = (valor10 - 25569) * 86400.0
                            xt10= datetime.datetime.utcfromtimestamp(seconds10).strftime('%d/%m/%Y')
                            nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                            self.tabla_base.setItem(i-1,j, nvalor10)
                        else:
                            valor = str(self.hoja_base.cell(i,j).value)
                            nvalor = QtWidgets.QTableWidgetItem(valor)
                            self.tabla_base.setItem(i-3,j, nvalor)
                            
                    else:
                        
                        valor = str(self.hoja_base.cell(i,j).value)
                        nvalor = QtWidgets.QTableWidgetItem(valor)
                        self.tabla_base.setItem(i-3,j, nvalor)
#                    valor = str(self.hoja_base.cell(i,j).value)
#                    nvalor = QtWidgets.QTableWidgetItem(valor)
#                    self.tabla_base.setItem(i-3,j, nvalor)
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
        for i in range(num_row):
            for j in range(num_col):
                if i == 0:
                    self.tabla_fact.setHorizontalHeaderItem(j, QtWidgets.QTableWidgetItem(str(self.hoja_facturacion.cell(i,j).value)))
                else: 
                    valor = str(self.hoja_facturacion.cell(i,j).value)
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
                            nvalor10 = QtWidgets.QTableWidgetItem(xt10)
                            self.tabla_sf.setItem(i-1,j, nvalor10)
                        else:
                            valor = str(self.hoja_sf.cell(i,j).value)
                            nvalor = QtWidgets.QTableWidgetItem(valor)
                            self.tabla_sf.setItem(i-1,j, nvalor)
#                    elif j==9:
#                        valor9 =self.hoja_sf.cell(i,j).value
#                        y = valor9 is float
#                        if y == True:
#                            seconds9 = (valor9 - 25569) * 86400.0
#                            xt9= datetime.datetime.utcfromtimestamp(seconds9).strftime('%d/%m/%Y')
#                            nvalor9 = QtWidgets.QTableWidgetItem(xt9)
#                            self.tabla_sf.setItem(i-1,j, nvalor9)
#                        else:
#                            valor = str(self.hoja_sf.cell(i,j).value)
#                            nvalor = QtWidgets.QTableWidgetItem(valor)
#                            self.tabla_sf.setItem(i-1,j, nvalor)
                    else:
                        
                        valor = str(self.hoja_sf.cell(i,j).value)
                        nvalor = QtWidgets.QTableWidgetItem(valor)
                        self.tabla_sf.setItem(i-1,j, nvalor)
        self.tabla_sf.resizeColumnsToContents()

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
        self.radio_regional.setText(_translate("MainWindow", "C&&W Networks Regional"))
        self.label_36.setText(_translate("MainWindow", "NRC"))
#        self.select_fact.setItemText(1, _translate("MainWindow", "Nuevos Servicios"))
#        self.select_fact.setItemText(2, _translate("MainWindow", "Novedades"))
#        self.select_fact.setItemText(3, _translate("MainWindow", "Renovaciones"))
        self.exportar_fact.setText(_translate("MainWindow", "Exportar"))
        self.label_2.setText(_translate("MainWindow", "Archivo Facturación"))
        self.label_17.setText(_translate("MainWindow", "NRC"))
        self.label_21.setText(_translate("MainWindow", "ID"))
        self.radio_colombia.setText(_translate("MainWindow", "C&&W Networks Colombia"))
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
        self.menuMenu.setTitle(_translate("MainWindow", "Menu"))
        self.menuCargar.setTitle(_translate("MainWindow", "Cargar"))
        self.actionCargar_Archivo_Base.setText(_translate("MainWindow", "Cargar Archivo Base"))
        self.actionCargar_Archivo_Facturacion.setText(_translate("MainWindow", "Cargar Archivo Facturación"))
        self.actionCargar_Archivo_Salesforce.setText(_translate("MainWindow", "Cargar Archivo Salesforce"))
        self.actionCambiar_Usuario.setText(_translate("MainWindow", "Cambiar Usuario"))
        self.actionExportar_todo.setText(_translate("MainWindow", "Exportar todo"))
        self.actionSalir.setText(_translate("MainWindow", "Salir"))
        self.actionManual.setText(_translate("MainWindow", "Manual"))
        self.label.setText(_translate("MainWindow", "TextLabel"))
        self.menuSettings.setTitle(_translate("MainWindow", "Settings"))
        self.menuMes.setTitle(_translate("MainWindow", "Mes"))
        self.menuRegion.setTitle(_translate("MainWindow", "Región"))
        self.actionEnero.setText(_translate("MainWindow", "Enero"))
        self.actionFebrero.setText(_translate("MainWindow", "Febrero"))
        self.actionAbril.setText(_translate("MainWindow", "Abril"))
        self.actionMayo.setText(_translate("MainWindow", "Mayo"))
        self.actionJunio.setText(_translate("MainWindow", "Junio"))
        self.actionJulio.setText(_translate("MainWindow", "Julio"))
        self.actionAgosto.setText(_translate("MainWindow", "Agosto"))
        self.actionSeptiembre.setText(_translate("MainWindow", "Septiembre"))
        self.actionOctubre.setText(_translate("MainWindow", "Octubre"))
        self.actionNoviembre.setText(_translate("MainWindow", "Noviembre"))
        self.actionDiciembre.setText(_translate("MainWindow", "Diciembre"))
        self.actionRegional.setText(_translate("MainWindow", "Regional"))
        self.actionColombia.setText(_translate("MainWindow", "Colombia"))
        self.actionMarzo.setText(_translate("MainWindow", "Marzo"))



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

