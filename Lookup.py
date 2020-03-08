import Ui_lookup
from PyQt5.QtWidgets import QDialog, QMessageBox

class Lookup(QDialog):
    def __init__(self):
        super().__init__()
        self.ui = Ui_lookup.Ui_Dialog()
        self.ui.setupUi(self)
