# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'c:\Users\nxf44756\OneDrive - NXP\Desktop\Gadgets\COTS\source\COTS\checkout.ui'
#
# Created by: PyQt5 UI code generator 5.13.0
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(1024, 768)
        Dialog.setMinimumSize(QtCore.QSize(1024, 768))
        Dialog.setMaximumSize(QtCore.QSize(1024, 768))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/Icon/resources/img/Checkout.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Dialog.setWindowIcon(icon)
        self.pushButton = QtWidgets.QPushButton(Dialog)
        self.pushButton.setGeometry(QtCore.QRect(860, 10, 151, 51))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton.setFont(font)
        self.pushButton.setFocusPolicy(QtCore.Qt.ClickFocus)
        self.pushButton.setObjectName("pushButton")

        self.retranslateUi(Dialog)
        self.pushButton.clicked.connect(Dialog.close)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "CQC Check-Out"))
        self.pushButton.setText(_translate("Dialog", "Back to Menu"))
import checkout_resource_rc
