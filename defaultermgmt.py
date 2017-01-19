# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'defaultermgmt.ui'
#
# Created by: PyQt5 UI code generator 5.6
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QFileDialog

import excelscrapper
import webbrowser
import os
import PyQt5

class Ui_MainWindow(object):

    def __init__(self):
        self.excelfilename = ""

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(597, 514)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        MainWindow.setSizePolicy(sizePolicy)
        MainWindow.setMinimumSize(QtCore.QSize(597, 514))
        MainWindow.setMaximumSize(QtCore.QSize(597, 514))
        font = QtGui.QFont()
        font.setPointSize(11)
        MainWindow.setFont(font)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.groupBox = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox.setGeometry(QtCore.QRect(310, 30, 261, 411))
        self.groupBox.setObjectName("groupBox")

        self.op_console = QtWidgets.QTextEdit(self.groupBox)
        self.op_console.setGeometry(QtCore.QRect(10, 30, 241, 371))
        self.op_console.setObjectName("op_console")

        self.send_button = QtWidgets.QPushButton(self.centralwidget)
        self.send_button.setGeometry(QtCore.QRect(60, 260, 161, 71))
        self.send_button.setObjectName("send_button")
        self.send_button.clicked.connect(self.start_process)

        self.get_names = QtWidgets.QPushButton(self.centralwidget)
        self.get_names.setGeometry(QtCore.QRect(60, 360, 161, 71))
        self.get_names.setObjectName("get_names")
        self.get_names.clicked.connect(self.defaulter_names)

        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setGeometry(QtCore.QRect(-10, 460, 611, 20))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")

        self.line_2 = QtWidgets.QFrame(self.centralwidget)
        self.line_2.setGeometry(QtCore.QRect(263, -10, 20, 481))
        self.line_2.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")

        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 450, 241, 16))
        font = QtGui.QFont()
        font.setPointSize(9)

        self.label.setFont(font)
        self.label.setObjectName("label")

        self.groupBox_2 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_2.setGeometry(QtCore.QRect(10, 20, 251, 51))
        self.groupBox_2.setObjectName("groupBox_2")

        self.addFile = QtWidgets.QPushButton(self.groupBox_2)
        self.addFile.setGeometry(QtCore.QRect(10, 20, 231, 23))
        self.addFile.setObjectName("addFile")
        self.addFile.clicked.connect(self.add_excel_file)

        self.filename = QtWidgets.QLabel(self.centralwidget)
        self.filename.setGeometry(QtCore.QRect(10, 120, 251, 16))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.filename.setFont(font)
        self.filename.setObjectName("filename")

        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Defaulter Management"))
        self.groupBox.setTitle(_translate("MainWindow", "Output Console"))
        self.send_button.setText(_translate("MainWindow", "Send"))
        self.get_names.setText(_translate("MainWindow", "Get Defaulter Names"))
        self.label.setText(_translate("MainWindow", "Created by Arun - Aditya - Sourabh"))
        self.groupBox_2.setTitle(_translate("MainWindow", "Excel File Name"))
        self.addFile.setText(_translate("MainWindow", "Add File"))
        self.filename.setText(_translate("MainWindow", "File Name:"))

    def start_process(self):
        self.op_console.append("Starting Process.....")
        if self.excelfilename == "":
            self.op_console.append("File Not Selected")
            self.op_console.append("Process Finished Unsuccessfully!")
        else:
            excelscrapper.run_defaulter_system(self.op_console, self.excelfilename)
            self.op_console.append("Process Finished Successfully!")

    def defaulter_names(self):
        self.op_console.append("Opening Notepad......")
        if os.path.exists('log.txt'):
            webbrowser.open('log.txt')
        else:
            self.op_console.append("File Not Found")

    def add_excel_file(self):
        dlg = QFileDialog()
        # dlg.setFilter("Excel File (*.xlsx)")
        dlg.setFileMode(QFileDialog.AnyFile)
        file_path = dlg.List
        if dlg.exec_():
            file_path = dlg.selectedFiles()
        file_name = str(file_path[0]).split("/")
        for x in file_name:
            if ".xlsx" in x:
                excel_name = x
        self.filename.setText(excel_name)
        self.excelfilename = file_path[0]


if __name__ == "__main__":
    import sys
    pyqt = os.path.dirname(PyQt5.__file__)
    QApplication.addLibraryPath(os.path.join(pyqt, "plugins"))
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

