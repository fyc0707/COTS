#coding=utf-8
import Ui_barcode, re
from io import BytesIO
from barcode import Code39
from barcode.writer import ImageWriter
from PyQt5.QtWidgets import QDialog, QErrorMessage
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import pyqtSlot

class Barcode(QDialog):
    def __init__(self):
        super().__init__()
        self.ui = Ui_barcode.Ui_Dialog()
        self.ui.setupUi(self)
        self.em = QErrorMessage(self)
        self.em.setWindowTitle('Error')
        
    
    @pyqtSlot()
    def on_barcodeGenButton_clicked(self):
        cqc_number = self.ui.cqcnumInput.text()
        if re.match(r'^[0-9]{6}[A-Z]{1}$', cqc_number):
            bc = BytesIO()
            Code39(cqc_number, ImageWriter(), False).write(bc, dict(text_distance=1.0, module_height=9, font_size=12))
            img = QPixmap.fromImage(QImage.fromData(bc.getvalue()))
<<<<<<< HEAD
            self.ui.barcodeDisplay.setPixmap(img.scaled(191, 81, 1))
=======
            self.ui.barcodeDisplay.setPixmap(img.scaled(241, 81, 1))            

>>>>>>> 6a2820d0bf2c0e6e8b8469e764c2f87b12bb49a9
        else:
            self.ui.cqcnumInput.setText('')
            self.em.showMessage('Please input a correct CQC Number')

