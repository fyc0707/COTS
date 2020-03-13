import sys
import CQCSniffer as cs
from PyQt5.QtWidgets import QApplication, QMainWindow
import Ui_Mainwindow, Receipt, Checkout, Barcode_w, Jerboa, Report, Lookup

        
class Mainwindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Mainwindow.Ui_MainWindow()
        self.ui.setupUi(self)

    def showWindow(self):
        self.hide()
        sender = self.sender().text()
        if sender == 'CQC Receipt':
            self.myDialog = Receipt.Receipt()
        elif sender == 'CQC Check-out   ':
            self.myDialog = Checkout.Checkout()
        elif sender == 'CQC Lookup':
            self.myDialog = Lookup.Lookup()
        elif sender == 'CQC WIP Report':
            self.myDialog = Report.Report()
        elif sender == 'JERBOA Queue':
            self.myDialog = Jerboa.Jerboa()
        elif sender == 'Barcode Scanner':
            self.myDialog = Barcode_w.Barcode()
        self.myDialog.exec_()
        self.show()

    def loginCQC(self):
        pass




if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = Mainwindow()
    w.show()
    sys.exit(app.exec_())