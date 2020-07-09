# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'filling_SV.ui'
#
# Created by: PyQt5 UI code generator 5.12.2
#
# WARNING! All changes made in this file will be lost!

import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog
import loading


class Ui_MainWindow(object):
    def __init__(self):
        super().__init__()

        self.mainwindow = QtWidgets.QMainWindow()
        self.setupUi(self.mainwindow)
        self.mainwindow.show()

    def setupUi(self, MainWindow):
        self.SV_FILENAME = ""
        self.COURSE = 0
        self.SESSION_FILENAME = ""
        self.SEMESTER = 0

        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(588, 443)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.l_choose_SV = QtWidgets.QLabel(self.centralwidget)
        self.l_choose_SV.setGeometry(QtCore.QRect(20, 10, 271, 71))
        self.l_choose_SV.setObjectName("l_choose_SV")
        self.pb_choose_SV = QtWidgets.QPushButton(self.centralwidget)
        self.pb_choose_SV.setGeometry(QtCore.QRect(330, 30, 151, 28))
        self.pb_choose_SV.setObjectName("pb_choose_SV")
        self.pb_choose_SV.clicked.connect(self.open_file_SV)
        self.l_enter_course = QtWidgets.QLabel(self.centralwidget)
        self.l_enter_course.setGeometry(QtCore.QRect(20, 110, 151, 71))
        self.l_enter_course.setObjectName("l_enter_course")
        self.l_choose_session = QtWidgets.QLabel(self.centralwidget)
        self.l_choose_session.setGeometry(QtCore.QRect(20, 180, 271, 71))
        self.l_choose_session.setObjectName("l_choose_session")
        self.le_enter_course = QtWidgets.QLineEdit(self.centralwidget)
        self.le_enter_course.setGeometry(QtCore.QRect(330, 130, 113, 22))
        self.le_enter_course.setObjectName("le_enter_course")
        self.pb_choose_session = QtWidgets.QPushButton(self.centralwidget)
        self.pb_choose_session.setGeometry(QtCore.QRect(330, 200, 151, 28))
        self.pb_choose_session.setObjectName("pb_choose_session")
        self.pb_choose_session.clicked.connect(self.open_file_SESSIONS)
        self.l_enter_semester = QtWidgets.QLabel(self.centralwidget)
        self.l_enter_semester.setGeometry(QtCore.QRect(20, 270, 311, 71))
        self.l_enter_semester.setObjectName("l_enter_semester")
        self.le_enter_semester = QtWidgets.QLineEdit(self.centralwidget)
        self.le_enter_semester.setGeometry(QtCore.QRect(330, 290, 113, 22))
        self.le_enter_semester.setObjectName("le_enter_semester")
        self.pb_start_filling = QtWidgets.QPushButton(self.centralwidget)
        self.pb_start_filling.setGeometry(QtCore.QRect(20, 340, 531, 28))
        self.pb_start_filling.setObjectName("pb_start_filling")
        self.pb_start_filling.clicked.connect(self.open_window)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.l_choose_SV.setText(_translate("MainWindow", "Выберите файл сводной ведомости на курс:"))
        self.pb_choose_SV.setText(_translate("MainWindow", "Выбрать файл"))
        self.l_enter_course.setText(_translate("MainWindow", "Введите номер курса:"))
        self.l_choose_session.setText(_translate("MainWindow", "Выберите файл файл сессии:"))
        self.pb_choose_session.setText(_translate("MainWindow", "Выбрать файл"))
        self.l_enter_semester.setText(_translate("MainWindow", "Введите номер семестра для выбранного курса:"))
        self.pb_start_filling.setText(_translate("MainWindow", "Перенести оценки в сводную ведомость"))

    def open_file_SV(self):
        filename = QFileDialog.getOpenFileName(self.mainwindow, "Открыть файл Excel", "C:\\Users\\gulya\\Desktop\\sessions\\documents")
        if filename[0]:
            self.SV_FILENAME = filename[0]
            #self.word = Document(self.FILENAME)
            #self.calculations()
        #print(self.SV_FILENAME)
        
    def open_file_SESSIONS(self):
        filename = QFileDialog.getOpenFileName(self.mainwindow, "Открыть файл Excel", "C:\\Users\\gulya\\Desktop\\sessions\\documents")
        if filename[0]:
            self.SESSIONS_FILENAME = filename[0]
            #self.word = Document(self.FILENAME)
            #self.calculations()
        #print(self.SESSIONS_FILENAME)

    def open_window(self):
        self.SV_FILENAME = "C:\\Users\\gulya\\Desktop\\sessions\\documents\\Сводная ведомость оценок.26курс.xlsx"
        self.SESSIONS_FILENAME = "C:\\Users\\gulya\\Desktop\\sessions\\documents\\сессия зима 16-17.xlsx"
        self.window = loading.Ui_MainWindow(self.SV_FILENAME, '26', self.SESSIONS_FILENAME, '2')
        #self.window = loading.Ui_MainWindow(self.SV_FILENAME, self.le_enter_course.text(), self.SESSIONS_FILENAME, self.le_enter_semester.text())
        self.mainwindow.close()

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = Ui_MainWindow()
    sys.exit(app.exec_())