import Ui_barcode
from PyQt5.QtWidgets import QDialog, QMessageBox

class Barcode(QDialog):
    def __init__(self):
        super().__init__()
        self.ui = Ui_barcode.Ui_Dialog()
        self.ui.setupUi(self)
