from PyQt5.QtWidgets import QDialog, QMessageBox

import Ui_manager
import CQCSniffer

class Manager(QDialog):
    def __init__(self, cs: CQCSniffer.CQCSniffer):
        super().__init__()
        self.ui = Ui_manager.Ui_Dialog()
        self.ui.setupUi(self)
        self.cs = cs
