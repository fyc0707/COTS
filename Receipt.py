import Ui_receipt
import CQCSniffer
from io import BytesIO
from barcode import Code39
from barcode.writer import ImageWriter
from PyQt5.QtWidgets import QDialog, QMessageBox
from PyQt5.QtCore import pyqtSlot

class Receipt(QDialog):
    def __init__(self, cs):
        super().__init__()
        self.ui = Ui_receipt.Ui_Dialog()
        self.ui.setupUi(self)
        self.cs = cs
        self.ui.welcomeLabel.setText('Welcome, '+self.cs.user_name)
    
    def itemSelected(self):
        pass

    def getCqcList(self):
        pass

    def checkin(self):
        pass

    #Reset the panel
    def reset(self):
        self.ui.cqcNumEdit.clear()
        self.ui.partNameEdit.clear()
        self.ui.cqeEdit.clear()
        self.ui.peEdit.clear()
        self.ui.rcvBox.setChecked(True)
        self.ui.prpBox.setChecked(True)
        self.ui.printBox.setChecked(True)
        self.ui.checkOnlyBox.setChecked(False)
        self.ui.progressBar.setValue(0)
        self.ui.resultLabel.setText(' ')

    #Logic for checkboxes
    @pyqtSlot()
    def on_rcvBox_clicked(self):
        if self.ui.rcvBox.isChecked():
            self.ui.checkOnlyBox.setChecked(False)
    @pyqtSlot()
    def on_prpBox_clicked(self):
        if self.ui.prpBox.isChecked():
            self.ui.checkOnlyBox.setChecked(False)
    @pyqtSlot()
    def on_printBox_clicked(self):
        if self.ui.printBox.isChecked():
            self.ui.checkOnlyBox.setChecked(False)
    @pyqtSlot()
    def on_checkOnlyBox_clicked(self):
        if self.ui.checkOnlyBox.isChecked():
            self.ui.rcvBox.setChecked(False)
            self.ui.prpBox.setChecked(False)
            self.ui.printBox.setChecked(False)

    def closeEvent(self, event):
        result = QMessageBox.question(self, "Message", "Confirm to exit. The unsubmitted job will be lost.", QMessageBox.Yes | QMessageBox.No)
        if(result == QMessageBox.Yes):
            event.accept()
        else:
            event.ignore()