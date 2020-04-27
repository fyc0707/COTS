#coding=utf-8
import os
import sys
from datetime import datetime

from PyQt5.QtWidgets import QApplication, QErrorMessage, QMainWindow, QMessageBox

import Barcode_w
import Checkout
import CQCSniffer
import Jerboa
import Lookup
import Receipt
import Report
import Ui_Mainwindow


class Mainwindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Mainwindow.Ui_MainWindow()
        self.ui.setupUi(self)
        self.em = QErrorMessage(self)
        self.em.setWindowTitle('Error')
        self.logged = False

    def showWindow(self):
        sender = self.sender().text()
        if sender == 'CQC Check-in' or sender == ' CQC Transfer    ':
            if self.logged:
                if self.cs.checkActive():
                    self.hide()
                    if sender == 'CQC Check-in':
                        self.myDialog = Receipt.Receipt(self.cs)
                    else:
                        self.myDialog = Lookup.Lookup(self.cs)
                    self.myDialog.exec_()
                    self.show()
                else:
                    self.em.showMessage('Session expired. Please log in.')
                    self.logged = False
                    self.ui.userName.show()
                    self.ui.password.show()
                    self.ui.loginButton.show()
                    self.ui.loginLabel.setText('WBI ID:\n\nPassword:')
            else:
                self.em.showMessage('Please log in.')
        else:
            self.hide()
            
            if sender == 'CQC Check-out':
                self.myDialog = Checkout.Checkout()
            elif sender == 'CQC WIP Report':
                self.myDialog = Report.Report()
            elif sender == 'JERBOA Queue':
                self.myDialog = Jerboa.Jerboa()
            elif sender == 'Barcode Scanner':
                self.myDialog = Barcode_w.Barcode()
            self.myDialog.exec_()
            self.show()

                        

    def loginCQC(self):
        self.cs = CQCSniffer.CQCSniffer('https://nww.cqc.nxp.com/CQC/', self.ui.userName.text(), self.ui.password.text())
        if self.cs.activeFlag:
            self.logged = True
            self.ui.userName.hide()
            self.ui.password.hide()
            self.ui.loginButton.hide()
            self.ui.loginLabel.setText('Welcome\n'+self.cs.user_name)
        else:
            self.em.showMessage('Failed to log in. Please check WBI acoount, password and intranet connection.')
            
    def closeEvent(self, event):
        result = QMessageBox.question(self, "Message", "Confirm to exit. Your account will be signed out.", QMessageBox.Yes | QMessageBox.No)
        if(result == QMessageBox.Yes):
            if self.logged:
                self.cs.logOut()
            event.accept()
        else:
            event.ignore()




if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = Mainwindow()
    w.show()
    log = 'log/'+datetime.today().date().isoformat()
    if os.path.exists(log):
        pass
    else:
        os.mkdir('log/'+datetime.today().date().isoformat())
    sys.exit(app.exec_())
