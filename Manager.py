import re
import os
from datetime import datetime

import pandas as pd
from PyQt5.QtCore import Qt, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtWidgets import QDialog, QErrorMessage, QHeaderView, QMessageBox

import Ui_manager
import CQCSniffer

class Manager(QDialog):
    def __init__(self, cs: CQCSniffer.CQCSniffer):
        super().__init__()
        self.ui = Ui_manager.Ui_Dialog()
        self.ui.setupUi(self)
        self.cs = cs
        self.em = QErrorMessage(self)
        self.partFlag = False
        self.checkFile()
        self.ui.engFuncEdit.addItems(['New role...', 'CQE', 'PE', 'TECHNICIAN'])
        self.ui.engFuncEdit.setCurrentIndex(-1)
        
    def closeEvent(self, event):
        if self.partFlag:
            result = QMessageBox.question(self, "Message", "Confirm to exit. The saved info will be lost.", QMessageBox.Yes | QMessageBox.No)
            if(result == QMessageBox.Yes):
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

    def checkFile(self):
        try:
            self.productTable = pd.read_csv('tables/ProductTable.csv', keep_default_na=False)
            self.ui.partNameEdit.addItem('Create a new product...')
            self.ui.partNameEdit.addItems(self.productTable['PART_TYPE_NAME'].sort_values().values.tolist())
            self.ui.partNameEdit.setCurrentIndex(-1)
        except Exception as err:
            self.em.showMessage('Failed to load the product table. Please close the file in use and restart the window.')
            print(err)
        try:
            self.engTable = pd.read_csv('tables/EmployeeTable.csv', keep_default_na=False)
            self.ui.engineerEdit.addItem('Create a new engineer...')
            self.ui.engineerEdit.addItems(self.engTable['NAME'].sort_values().values.tolist())
            self.ui.engineerEdit.setCurrentIndex(-1)
        except Exception as err:
            self.em.showMessage('Failed to load the engineer table. Please close the file in use and restart the window.')
            print(err)

    def engineerSearch(self):
        self.ui.engResultLabel.setText('')
        self.ui.engFuncEdit.setCurrentIndex(-1)
        self.ui.engInfoLabel.setText('')
        self.ui.engEmailLabel.setText('')
        self.ui.mgrNameLabel.setText('')
        self.ui.mgrEmailLabel.setText('')
        self.ui.engAttEdit.setPlainText('')
        eng = self.ui.engineerEdit.currentText()
        if eng in self.engTable['NAME'].values:
            row = self.engTable[self.engTable['NAME']==eng].iloc[0]
            self.ui.engInfoLabel.setText(row['NAME'])
            self.ui.engEmailLabel.setText(row['EMAIL'])
            self.ui.engFuncEdit.setCurrentText(row['FUNCTION'])
            self.ui.mgrNameLabel.setText(row['MANAGER'])
            self.ui.mgrEmailLabel.setText(row['MANAGER_EMAIL'])
            att = str(row['ATTENTION_NAME'])
            att = att.replace(';',';\n', att.count(';')-1)
            self.ui.engAttEdit.setPlainText(att)
        else:
            if re.search(r'[A-Za-z]{3}[0-9]{5}',eng):
                self.thread = fillInfoThread(self.cs, eng)
                self.thread.result_signal.connect(self.fillInfoCallBack)
                self.thread.start()
                self.busy()
            else:
                self.em.showMessage('Please input valid WBI account.')
                
    def fillInfoCallBack(self, signal):
        if signal==101:
            self.em.showMessage('CQC system handling error. Please contact COTS developer.')
            self.release()
            self.thread.exit()
        elif signal==103:
            self.em.showMessage('Session expired. Please restart the application.')
            self.release()
            self.thread.exit()
        else:
            self.release()
            self.ui.engInfoLabel.setText(self.thread.name)
            if self.thread.name in self.engTable['NAME'].values:
                self.ui.engFuncEdit.setCurrentText(self.engTable[self.engTable['NAME']==self.thread.name]['FUNCTION'].iloc[0])
                att = str(self.engTable[self.engTable['NAME']==self.thread.name]['ATTENTION_NAME'].iloc[0])
                att = att.replace(';',';\n', att.count(';')-1)
                self.ui.engAttEdit.setPlainText(att)
            self.ui.engEmailLabel.setText(self.thread.email)
            self.ui.mgrNameLabel.setText(self.thread.mgr)
            self.ui.mgrEmailLabel.setText(self.thread.mgr_email)
            self.thread.exit()
            if self.ui.engineerEdit.isEditable():
                self.ui.engineerEdit.setCurrentIndex(-1)

    def openCSV(self):
        sender = self.sender().objectName()
        if sender == 'partCSVButton':  
            try:
                os.startfile(os.path.abspath('tables/ProductTable.csv'))
            except Exception as err:
                self.em.showMessage('Failed to open the file. Please close the file in use.')
                print(err)
        elif sender == 'engCSVButton':
            try:
                os.startfile(os.path.abspath('tables/EmployeeTable.csv'))
            except Exception as err:
                self.em.showMessage('Failed to open the file. Please close the file in use.')
                print(err)
    
    def partSelected(self):
        self.partFlag = False
        self.ui.respResultLabel.setText('')
        self.ui.peOwnerLabel.setText('')
        self.ui.partAttEdit.setPlainText('')
        try:
            part = self.ui.partNameEdit.currentText()
            if part == '':
                return
            self.ui.partLabel.setText('"' + part + '"')
            if part in self.productTable['PART_TYPE_NAME'].values:
                self.ui.peOwnerLabel.setText(self.productTable[self.productTable['PART_TYPE_NAME']==part]['PE_NAME'].iloc[0])
                att = str(self.productTable[self.productTable['PART_TYPE_NAME']==part]['ATTENTION_NAME'].iloc[0])
                att = att.replace(';', ';\n', att.count(';')-1)
                self.ui.partAttEdit.setPlainText(att)
            else:
                pass
            self.partFlag = True
        except Exception as err:
            print(err)

    def saveResp(self):
        if not self.partFlag:
            return
        part = self.ui.partLabel.text()[1:-1]
        if part in self.productTable['PART_TYPE_NAME'].values:
            self.productTable.loc[self.productTable['PART_TYPE_NAME']==part] = [[part, '', self.ui.peOwnerLabel.text(), self.ui.partAttEdit.toPlainText().replace('\n','')]]
        else:
            self.productTable.loc[len(self.productTable)] = [part, '', self.ui.peOwnerLabel.text(), self.ui.partAttEdit.toPlainText().replace('\n','')]
            self.ui.partNameEdit.addItem(part)
        try:
            self.productTable.to_csv('tables/ProductTable.csv', index_label=False, index=False)
            self.ui.respResultLabel.setText('Successfully updated the table')
        except Exception as err:
            print(err)
            self.em.showMessage('Failed to update the product table. Please close the file in use and retry.')
            self.ui.respResultLabel.setText('')
        
    def saveEng(self):
        eng = self.ui.engInfoLabel.text()
        if eng == '':
            return
        if eng in self.engTable['NAME'].values:
            self.engTable.loc[self.engTable['NAME']==eng] = [[self.ui.engInfoLabel.text(), self.ui.engEmailLabel.text(), 
                                                                self.ui.engFuncEdit.currentText(), self.ui.mgrNameLabel.text(), self.ui.mgrEmailLabel.text(), self.ui.engAttEdit.toPlainText().replace('\n','')]]
        else:
            self.engTable.loc[len(self.engTable)] = [self.ui.engInfoLabel.text(), self.ui.engEmailLabel.text(), 
                                                    self.ui.engFuncEdit.currentText(), self.ui.mgrNameLabel.text(), self.ui.mgrEmailLabel.text(), self.ui.engAttEdit.toPlainText().replace('\n','')]
            self.ui.engineerEdit.addItem(eng)
        try:
            self.engTable.to_csv('tables/EmployeeTable.csv', index_label=False, index=False)
            self.ui.engResultLabel.setText('Successfully updated the table')
        except Exception as err:
            print(err)
            self.em.showMessage('Failed to update the engineer table. Please close the file in use and retry.')
            self.ui.engResultLabel.setText('')

    def clearEdit(self):
        sender = self.sender().objectName()
        if sender == 'clearOwnerButton':
            self.ui.peOwnerLabel.clear()
            self.ui.respResultLabel.setText('')
        elif sender == 'clearAttButton':
            self.ui.partAttEdit.clear()
            self.ui.respResultLabel.setText('')
        elif sender == 'clearEngAttButton':
            self.ui.engAttEdit.clear()
            self.ui.engResultLabel.setText('')

    def assignPE(self):
        self.ui.respResultLabel.setText('')
        if self.ui.engineerEdit.isEditable == '' or (not self.partFlag):
            return
        self.ui.peOwnerLabel.setText(self.ui.engineerEdit.currentText())

    def addPartAttention(self):
        eng = self.ui.engineerEdit.currentText()
        self.ui.respResultLabel.setText('')
        text = self.ui.partAttEdit.toPlainText()
        if eng == '' or self.ui.engineerEdit.isEditable():
            return
        if eng in text.replace('\n','').split(';'):
            pass
        else:
            if text != '':
                if text.replace('\n','')[-1] != ';':
                    self.ui.partAttEdit.setPlainText(text+';')
            self.ui.partAttEdit.appendPlainText(eng+';')

    def addEngAttention(self):
        eng = self.ui.engineerEdit.currentText()
        text = self.ui.engAttEdit.toPlainText()
        if eng == '' or self.ui.engineerEdit.isEditable():
            return
        if eng in text.replace('\n','').split(';'):
            pass
        else:
            if text != '':
                if text.replace('\n','')[-1] != ';':
                    self.ui.partAttEdit.setPlainText(text+';')
            self.ui.engAttEdit.appendPlainText(eng+';')

    def updateEng(self):
        self.ui.engResultLabel.setText('')
        eng = self.ui.engInfoLabel.currentText()
        if eng == '':
            return
        self.thread = fillInfoThread(self.cs, eng)
        self.thread.result_signal.connect(self.fillInfoCallBack)
        self.thread.start()
        self.busy()
    
    def updateAllEng(self):
        self.ui.engResultLabel.setText('')

    def indexChanged(self):
        sender = self.sender()
        index = sender.currentIndex()
        text = sender.currentText()
        if (index == 0) and ('Create' in text or 'New' in text):
            sender.setEditable(True)
            sender.setEditText('')
        elif index == -1:
            sender.setEditable(False)
        else:
            sender.setEditable(False)
        
    def busy(self):
        self.ui.lookUpButton.setEnabled(False)
        self.ui.updateAllButton.setEnabled(False)
        self.ui.updateButton.setEnabled(False)

    def release(self):
        self.ui.lookUpButton.setEnabled(True)
        self.ui.updateAllButton.setEnabled(True)
        self.ui.updateButton.setEnabled(True)


class fillInfoThread(QThread):
    
    result_signal = pyqtSignal(int)
    
    def __init__(self, cs: CQCSniffer.CQCSniffer, wbi):
        super(fillInfoThread, self).__init__()
        self.cs = cs
        self.wbi = wbi
        self.name = ''
        self.email = ''
        self.mgr = ''
        self.mgr_email = ''

    def run(self):
        try:
            if self.cs.checkActive():
                self.name, self.email, self.mgr, self.mgr_email = self.cs.getFullInfo(self.wbi)
                self.result_signal.emit(102)
            else:
                self.result_signal.emit(103)
        except Exception as err:
            print(err)
            self.result_signal.emit(101)