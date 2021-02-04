#coding=utf-8
import re



from PyQt5.QtCore import pyqtSlot
from PyQt5.QtGui import QImage, QPixmap
from PyQt5.QtWidgets import QDialog, QErrorMessage

import Ui_shipment
import CQCSniffer

class Shipment(QDialog):
    def __init__(self, cs: CQCSniffer.CQCSniffer):
        super().__init__()
        self.ui = Ui_shipment.Ui_Dialog()
        self.ui.setupUi(self)
        self.em = QErrorMessage(self)
        self.em.setWindowTitle('Error')
        self.cs = cs
        
    
