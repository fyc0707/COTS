import Ui_checkout
import CQCSniffer
from PyQt5.QtWidgets import QDialog, QMessageBox

class Checkout(QDialog):
    def __init__(self):
        super().__init__()
        self.ui = Ui_checkout.Ui_Dialog()
        self.ui.setupUi(self)
        
    
    