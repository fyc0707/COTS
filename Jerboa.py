import Ui_jerboa
from PyQt5.QtWidgets import QDialog, QMessageBox

class Jerboa(QDialog):
    def __init__(self):
        super().__init__()
        self.ui = Ui_jerboa.Ui_Dialog()
        self.ui.setupUi(self)
