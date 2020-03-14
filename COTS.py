import sys, os
import CQCSniffer
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QErrorMessage
import Ui_Mainwindow, Receipt, Checkout, Barcode_w, Jerboa, Report, Lookup

        
class Mainwindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Mainwindow.Ui_MainWindow()
        self.ui.setupUi(self)
        self.em = QErrorMessage(self)
        self.em.setWindowTitle('Error')
        self.logged = False

    def showWindow(self):
        if self.logged:
            if self.cs.checkActive():
                self.hide()
                sender = self.sender().text()
                if sender == 'CQC Check-in':
                    self.myDialog = Receipt.Receipt(self.cs)
                elif sender == 'CQC Check-out   ':
                    self.myDialog = Checkout.Checkout(self.cs)
                elif sender == 'CQC Lookup':
                    self.myDialog = Lookup.Lookup()
                elif sender == 'CQC WIP Report':
                    self.myDialog = Report.Report()
                elif sender == 'JERBOA Queue':
                    self.myDialog = Jerboa.Jerboa()
                elif sender == 'Barcode Scanner':
                    self.myDialog = Barcode_w.Barcode()
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
            event.accept()
        else:
            event.ignore()




if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = Mainwindow()
    w.show()
    try:
        os.mkdir('log/'+datetime.today().date().isoformat())
    except:
        pass
    sys.exit(app.exec_())