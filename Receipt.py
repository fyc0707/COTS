import Ui_receipt
from PyQt5.QtWidgets import QDialog, QMessageBox

class Receipt(QDialog):
    def __init__(self):
        super().__init__()
        self.ui = Ui_receipt.Ui_Dialog()
        self.ui.setupUi(self)
    
    def closeEvent(self, event):
        result = QMessageBox.question(self, "Message", "Confirm to exit. The unsubmitted job will be lost.", QMessageBox.Yes | QMessageBox.No)
        if(result == QMessageBox.Yes):
            event.accept()
        else:
            event.ignore()