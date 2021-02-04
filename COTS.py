#coding=utf-8
import os
import sys
import re
import pandas as pd 
from datetime import datetime

from PyQt5.QtWidgets import QApplication, QErrorMessage, QMainWindow, QMessageBox

import Shipment
import Checkout
import CQCSniffer
import Manager
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
        self.cs = CQCSniffer.CQCSniffer('https://nww.cqc.nxp.com/CQC/', '', '')

    def showWindow(self):
        sender = self.sender().text()
        if self.logged:
            if sender == 'CQC Check-out':
                self.hide()
                self.myDialog = Checkout.Checkout()
                self.myDialog.exec_()
                self.show()
                return
            if self.cs.checkActive():
                self.hide()
                self.myDialog = self.setDiag(sender)
                self.myDialog.exec_()
                self.show()
            else:
                result = QMessageBox.question(self, 'Message', 'The session connected to CQC system is expired. Continue with offline mode?', QMessageBox.Yes | QMessageBox.No)
                if result == QMessageBox.Yes:
                    self.hide()
                    self.myDialog = self.setDiag(sender)
                    self.myDialog.exec_()
                    self.show()
                else:
                    self.logged = False
                    self.ui.userName.show()
                    self.ui.password.show()
                    self.ui.loginButton.show()
                    self.ui.loginLabel.setText('WBI ID:\n\nPassword:')
        else:
            result = QMessageBox.question(self, 'Message', 'Not logged on CQC system. Continue with offline mode?', QMessageBox.Yes | QMessageBox.No)
            if result == QMessageBox.Yes:
                self.hide()
                self.myDialog = self.setDiag(sender)
                self.myDialog.exec_()
                self.show()
    
    def setDiag(self, sender):
        if sender == 'CQC Check-out':
            return Checkout.Checkout()
        elif sender == 'CQC WIP Report':
            return Report.Report(self.cs)
        elif sender == 'Product Manager':
            return Manager.Manager(self.cs)
        elif sender == 'CQC on the Way':
            return Shipment.Shipment(self.cs)
        elif sender == 'CQC Check-in':
            return Receipt.Receipt(self.cs)
        elif sender == ' CQC Transfer    ':
            return Lookup.Lookup(self.cs)

    def loginCQC(self):
        self.cs = CQCSniffer.CQCSniffer('https://nww.cqc.nxp.com/CQC/', self.ui.userName.text(), self.ui.password.text())
        self.cs.login()
        if self.cs.activeFlag:
            self.logged = True
            self.ui.userName.hide()
            self.ui.password.hide()
            self.ui.loginButton.hide()
            self.ui.loginLabel.setText('Welcome\n'+self.cs.user_name)
        else:
            self.em.showMessage('Failed to log in. Please check WBI acoount, password and intranet connection.')
            
    def closeEvent(self, event):
        if self.logged:
            result = QMessageBox.question(self, "Message", "Confirm to exit. Your account will be signed out.", QMessageBox.Yes | QMessageBox.No)
            if(result == QMessageBox.Yes):
                self.cs.logOut()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()




if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = Mainwindow()
    w.show()
    
    history = os.listdir('log/')
    lastlog = None
    for h in history:
        if re.match(r'^\d\d\d\d-\d\d-\d\d$', h) and os.path.exists('log/'+h+'/log.csv'):
            time = datetime.strptime(h, '%Y-%m-%d')
            if lastlog == None:
                lastlog = time
            else:
                if time > lastlog:
                    lastlog = time
    lastlog = lastlog.date().isoformat() if lastlog != None else None
    log = 'log/'+datetime.today().date().isoformat()
    if os.path.exists(log):
        pass
    else:
        if lastlog != None:
            try:
                leftover = pd.read_csv('log/'+lastlog+'/log.csv', keep_default_na=False)
                leftover = leftover[(leftover['Checkout']=='') & (leftover['Status']!='S')]
                for i in leftover.index.to_list():
                    leftover.loc[i,'Checkin'] = 'N'
                os.mkdir(log)
                leftover.to_csv(log+'/log.csv', index_label=False, index=False)
            except Exception as err:
                print(err)
                w.em.showMessage('Failed to acquire the log of yesterday. The program will exit. Please close the files in use and retry.')
                w.close()
        else:
            os.mkdir(log)

    
    sys.exit(app.exec_())