import Ui_checkout
from PyQt5.QtWidgets import QDialog, QMessageBox

class Checkout(QDialog):
    def __init__(self):
        super().__init__()
        self.ui = Ui_checkout.Ui_Dialog()
        self.ui.setupUi(self)
    
    def closeEvent(self, event):
        result = QMessageBox.question(self, "Message", "Confirm to exit. The unsubmitted job will be lost.", QMessageBox.Yes | QMessageBox.No)
        if(result == QMessageBox.Yes):
            event.accept()
        else:
            event.ignore()