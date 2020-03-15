#coding=utf-8
import os
from datetime import datetime
from io import BytesIO

import pandas as pd
from barcode import Code39
from barcode.writer import ImageWriter
from PyQt5.QtCore import QAbstractTableModel, Qt, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtWidgets import QDialog, QErrorMessage, QMessageBox

import CQCSniffer
import Ui_receipt


class Receipt(QDialog):
    def __init__(self, cs):
        super().__init__()
        self.ui = Ui_receipt.Ui_Dialog()
        self.ui.setupUi(self)
        self.cs = cs
        self.em = QErrorMessage(self)
        self.em.setWindowTitle('Error')
        self.ui.welcomeLabel.setText('Welcome, '+self.cs.user_name)
        self.list_file = 'log/'+datetime.today().date().isoformat()+'/rcvList.xlsx'
        self.checkFile()
        
    
    def itemSelected(self):
        row = self.ui.cqcList.selectedIndexes()[0].row()
        self.ui.cqcNumEdit.setText(self.rcv_df['CQC#'].loc[row])
        self.ui.partNameEdit.setText(self.rcv_df['Part Name'].loc[row])
        self.ui.cqeEdit.setText(self.rcv_df['CQE'].loc[row])
        #self.ui.peEdit.setText(self.rcv_df['PE'].loc[row])
        self.ui.instruEdit.setText(self.rcv_df['Instruction'].loc[row])


    def getCqcList(self):
        self.reset()
        fileobj = open(self.list_file+'1', 'wb')
        self.thread = downloadThread(self.cs, fileobj, 1024)
        self.thread.process_signal.connect(self.downloadCallBack)
        self.thread.start()
        self.busy()
        

    def downloadCallBack(self, signal):
        if signal == 101:
            self.release()
            self.checkFile()
            self.em.showMessage('Download failed. Please retry or restart the application.')
            os.remove(self.list_file+'1')
        elif signal == 102:
            self.release()
            self.ui.resultLabel.setText('Download successfully')
            os.replace(self.list_file+'1', self.list_file)
            self.checkFile()
        else:
            self.ui.progressBar.setValue(signal)

    def checkin(self):
        pass
    
    def checkinCallBack(self, signal):
        pass


    def checkFile(self):
        if os.path.exists(self.list_file):
            self.ui.cqcListLable.setText('Last Update:\n'+datetime.fromtimestamp(os.path.getmtime(self.list_file)).strftime('%Y-%m-%d %H:%M:%S'))
            self.rcv_df = self.cs.getRCVData(self.list_file)
            model = pandasModel(self.rcv_df.drop(['B2B'], axis=1))
            self.ui.cqcList.setModel(model)
            self.ui.cqcList.setColumnWidth(0,55)
            self.ui.cqcList.setColumnWidth(1,35)
            self.ui.cqcList.setColumnWidth(2,90)
            self.ui.cqcList.setColumnWidth(3,85)
            self.ui.cqcList.setColumnWidth(4,85)
            self.ui.cqcList.setColumnWidth(5,90)
            self.ui.cqcList.setColumnWidth(6,190)

        else:
            self.ui.cqcListLable.setText('No list found')
    

    def reset(self):
        '''Reset the panel
        '''
        self.ui.cqcNumEdit.clear()
        self.ui.partNameEdit.clear()
        self.ui.cqeEdit.clear()
        self.ui.peEdit.clear()
        self.ui.instruEdit.clear()
        self.ui.rcvBox.setChecked(True)
        self.ui.prpBox.setChecked(True)
        self.ui.printBox.setChecked(True)
        self.ui.checkOnlyBox.setChecked(False)
        self.ui.progressBar.setValue(0)
        self.ui.resultLabel.setText(' ')

    def busy(self):
        self.ui.getListButton.setEnabled(False)
        self.ui.checkinButton.setEnabled(False)

    def release(self):
        self.ui.getListButton.setEnabled(True)
        self.ui.checkinButton.setEnabled(True)
  

    @pyqtSlot()
    def on_rcvBox_clicked(self):
        if self.ui.rcvBox.isChecked():
            self.ui.checkOnlyBox.setChecked(False)
    @pyqtSlot()
    def on_prpBox_clicked(self):
        if self.ui.prpBox.isChecked():
            self.ui.checkOnlyBox.setChecked(False)
    @pyqtSlot()
    def on_printBox_clicked(self):
        if self.ui.printBox.isChecked():
            self.ui.checkOnlyBox.setChecked(False)
    @pyqtSlot()
    def on_checkOnlyBox_clicked(self):
        if self.ui.checkOnlyBox.isChecked():
            self.ui.rcvBox.setChecked(False)
            self.ui.prpBox.setChecked(False)
            self.ui.printBox.setChecked(False)


    def closeEvent(self, event):
        result = QMessageBox.question(self, "Message", "Confirm to exit. The unsubmitted job will be lost.", QMessageBox.Yes | QMessageBox.No)
        if(result == QMessageBox.Yes):
            event.accept()
        else:
            event.ignore()


class pandasModel(QAbstractTableModel):
    '''Model dataframe to QTableView
    '''
    def __init__(self, data):
        super(pandasModel, self).__init__()
        self.data = data
    
    def rowCount(self, parent=None):
        return self.data.shape[0]
    
    def columnCount(self, parent=None):
        return self.data.shape[1]
    
    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self.data.iloc[index.row(), index.column()])
        return None
    
    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.data.columns[section]
        return None
    

class downloadThread(QThread):
    '''RCV file download thread
    '''
    process_signal = pyqtSignal(int)

    def __init__(self, cs, fileobj, buffer):
        super(downloadThread, self).__init__()
        self.cs = cs
        self.filesize = None
        self.fileobj = fileobj
        self.buffer = buffer

    def run(self):
        try:
            if self.cs.checkActive():
                self.cs.session.get(self.cs.url+'advancedSearch.do?method=advancedSearchBookMarkResults&bookid=8632', headers=self.cs.headers, verify=False, timeout=700)
                f = self.cs.session.get(self.cs.url+'advancedSearch.do?method=advancedSearchResultsExcelExport', headers=self.cs.headers, verify=False, timeout=700, stream=True)
                size = len(f.content)
                offset = 0
                for chunk in f.iter_content(chunk_size=self.buffer):
                    if not chunk:
                        break
                    self.fileobj.seek(offset)
                    self.fileobj.write(chunk)
                    offset = offset + len(chunk)
                    process = offset / int(size) * 100
                    self.process_signal.emit(int(process))
                self.process_signal.emit(102)
            else:
                self.process_signal.emit(101)
        except:
            self.process_signal.emit(101)
        self.fileobj.close()
        self.exit(0)
            


class checkinThread(QThread):
    '''Checkin job chain thread
    '''
    process_signal = pyqtSignal(int)
    success_signal = pyqtSignal(str)

    def __init__(self, mode):
        super(checkinThread, self).__init__()
    
    def run(self):
        pass
