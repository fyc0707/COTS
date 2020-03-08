import sys
from PyQt5.QtWidgets import QApplication, QWidget, QMainWindow
import Ui_Mainwindow

if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = QMainWindow()
    ui = Ui_Mainwindow.Ui_MainWindow()
    ui.setupUi(w)
    w.show()
    sys.exit(app.exec_())