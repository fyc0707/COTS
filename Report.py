import Ui_report
from PyQt5.QtWidgets import QDialog, QMessageBox

class Report(QDialog):
    def __init__(self):
        super().__init__()
        self.ui = Ui_report.Ui_Dialog()
        self.ui.setupUi(self)
