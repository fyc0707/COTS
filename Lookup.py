import Ui_lookup
from PyQt5.QtWidgets import QDialog, QMessageBox

class Lookup(QDialog):
    def __init__(self, cs):
        super().__init__()
        self.ui = Ui_lookup.Ui_Dialog()
        self.ui.setupUi(self)
        self.cs = cs

    def closeEvent(self, event):
        result = QMessageBox.question(self, "Message", "Confirm to exit. The unsubmitted job will be lost.", QMessageBox.Yes | QMessageBox.No)
        if(result == QMessageBox.Yes):
            event.accept()
        else:
            event.ignore()