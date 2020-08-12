import os
import re
from datetime import datetime

import pandas as pd
import pythoncom
pythoncom.CoInitialize()
from PyQt5.QtCore import QAbstractTableModel, Qt, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtWidgets import QCompleter, QDialog, QErrorMessage, QMessageBox
from win32com.client import Dispatch
from HTMLTable import HTMLTable

import CQCSniffer
import Ui_lookup


class Lookup(QDialog):
    def __init__(self, cs: CQCSniffer.CQCSniffer):
        super().__init__()
        self.ui = Ui_lookup.Ui_Dialog()
        self.ui.setupUi(self)
        self.cs = cs
        if not self.cs.activeFlag:
            self.ui.prpBox.setCheckable(False)
            self.ui.tstBox.setCheckable(False)
        
        self.em = QErrorMessage(self)
        self.em.setWindowTitle('Error')
        self.ui.welcomeLabel.setText('Welcome, ' + self.cs.user_name)
        self.list_file = 'log/'+datetime.today().date().isoformat()+'/wipList.xlsx'
        self.log_file = 'log/'+datetime.today().date().isoformat()+'/log.csv'
        self.data = None
        self.queue = pd.DataFrame(columns=['CQC#', 'CQE', 'PE', 'Product', 'TST', 'Instruction'])
        self.checkFile()
        

    def transfer(self):
        try:
            self.ui.resultLabel.setText('')
            if (self.ui.prpBox.isChecked() or self.ui.tstBox.isChecked()):
                if re.match(r'^[0-9]{6}[A-Z]{1}$', self.ui.cqcNumEdit.text()):
                    cqc_num = self.ui.cqcNumEdit.text()
                else:
                    self.em.showMessage('Bad CQC number.')
                    return
                cqe = self.ui.cqeEdit.text()
                pe = self.ui.peEdit.text()
                if (cqe == None or pe == None):
                    self.em.showMessage('Please input CQE and PE information')
                    return
                product = self.ui.partNameEdit.text()
                ins = self.ui.insEdit.text()
                tst_flag = False
                cqc_info = [cqc_num, cqe, pe, product]
                mode = [self.ui.prpBox.isChecked(), self.ui.tstBox.isChecked()]
                self.queue.loc[len(self.queue)] = [cqc_num, cqe, pe, product, False, ins]
                self.queue.drop_duplicates(['CQC#'], keep='last', ignore_index=True, inplace=True)
                self.thread = transferThread(self.cs, mode, cqc_info)
                self.thread.result_signal.connect(self.transferCallBack)
                self.thread.start()
                self.busy()
            else:
                self.queue.loc[len(self.queue)] = [self.ui.cqcNumEdit.text(), self.ui.cqeEdit.text(), self.ui.peEdit.text(), self.ui.partNameEdit.text(), False, self.ui.insEdit.text()]
                self.queue.drop_duplicates(['CQC#'], keep='last', ignore_index=True, inplace=True)
                self.reset()
                self.updateTable()                
        except Exception as err:
            print(err)  
            
    def transferCallBack(self, signal):
        if signal=='101':
            self.em.showMessage('CQC system handling error. Please contact COTS admin.')
            self.reset()
            self.release()
            self.updateTable()
        elif signal=='103':
            self.em.showMessage('Session expired. Please restart the application.')
            self.reset()
            self.release()
            self.updateTable()
        elif signal=='100':
            self.queue.iloc[-1]['TST'] = self.thread.tst_flag
            self.reset()
            self.release()
            self.updateTable()
        else:
            self.ui.resultLabel.setText(self.ui.resultLabel.text()+signal)
    
    def updateTable(self):
        model = pandasModel(self.queue)
        self.ui.cqcList.setModel(model)
        self.ui.cqcList.setColumnWidth(0,90)
        self.ui.cqcList.setColumnWidth(1,120)
        self.ui.cqcList.setColumnWidth(2,120)
        self.ui.cqcList.setColumnWidth(3,100)
        self.ui.cqcList.setColumnWidth(4,40)
        self.ui.cqcList.setColumnWidth(5,100)

    def itemSelected(self):
        index = self.ui.cqcList.selectedIndexes()[0].row()
        ans = QMessageBox.question(self, 'Message', 'Confirm to remove CQC '+self.queue.iloc[index]['CQC#'])
        if ans == QMessageBox.Yes:
            self.queue.drop(axis=0, index=index, inplace=True)
            self.queue.reset_index(drop=True, inplace=True)
        self.updateTable()

    def email(self):
        if len(self.queue)==0:
            self.em.showMessage('The queue is empty.')
            return
        self.thread = emailThread(self.cs, self.queue, self.cqeTable, self.peTable)
        self.thread.result_signal.connect(self.emailCallBack)
        self.thread.start()
        self.busy()
    
    def emailCallBack(self, signal):
        if signal == '100':
            self.ui.resultLabel.setText('Email sent successfully.')
            pythoncom.CoUninitialize()
            self.release()
        else:
            self.ui.resultLabel.setText('Email failed.')
            pythoncom.CoUninitialize()
            self.release()


    @pyqtSlot()
    def on_clearButton_clicked(self):
        self.reset()
        self.ui.resultLabel.setText('')
        self.data = None

    @pyqtSlot()
    def on_listenerButton_clicked(self):
        self.data = self.ui.cqcNumEdit.text()
        self.reset()
        self.ui.resultLabel.setText('')
        try:
            if '（' in self.data or '）' in self.data:
                self.em.showMessage('Please change the input language to English.')
                self.data = None
            else:
                self.data = self.data.split('/\\')
                if len(self.data)==12:
                    cqc_num, qty, code, ship, cqe, pe, pem, part_name, ins, rcv, prp, time = self.data
                    self.ui.cqcNumEdit.setText(cqc_num)
                    self.ui.partNameEdit.setText(part_name)
                    self.ui.cqeEdit.setText(cqe)
                    self.ui.peEdit.setText(pe)
                    self.ui.insEdit.setText(ins)
                    self.ui.transferButton.setFocus()
                else:
                    self.data = None
                    self.em.showMessage('Unidentified QR code.')
        except Exception as err:
            self.data = None
            print(err)

    def closeEvent(self, event):
        result = QMessageBox.question(self, "Message", "Confirm to exit. The unsubmitted job will be lost.", QMessageBox.Yes | QMessageBox.No)
        if(result == QMessageBox.Yes):
            event.accept()
        else:
            event.ignore()

    def checkFile(self):
        try:
            self.productTable = pd.read_csv('tables/ProductTable.csv', keep_default_na=False)
            completer = QCompleter(self.productTable['PART_TYPE_NAME'].values.tolist())
            completer.setFilterMode(Qt.MatchContains)
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            self.ui.partNameEdit.setCompleter(completer)
        except Exception as err:
            self.em.showMessage('Failed to load the product table. Please close the file in use and restart the window.')
            print(err)
        try:
            self.peTable = pd.read_csv('tables/PETable.csv', keep_default_na=False)
            completer = QCompleter(self.peTable['PE_NAME'].values.tolist())
            completer.setFilterMode(Qt.MatchContains)
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            self.ui.peEdit.setCompleter(completer)
        except Exception as err:
            self.em.showMessage('Failed to load the PE table. Please close the file in use and restart the window.')
            print(err)
        try:
            self.cqeTable = pd.read_csv('tables/CQETable.csv', keep_default_na=False)
            completer = QCompleter(self.cqeTable['CQE_NAME'].values.tolist())
            completer.setFilterMode(Qt.MatchContains)
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            self.ui.cqeEdit.setCompleter(completer)
        except Exception as err:
            self.em.showMessage('Failed to load the CQE table. Please close the file in use and restart the window.')
            print(err)
        


    def reset(self):
        self.ui.cqcNumEdit.clear()
        self.ui.partNameEdit.clear()
        self.ui.cqeEdit.clear()
        self.ui.peEdit.clear()
        self.ui.insEdit.clear()
        if self.ui.tstBox.isCheckable():
            self.ui.tstBox.setChecked(True)
            self.ui.prpBox.setChecked(True)
        else:
            self.ui.tstBox.setChecked(False)
            self.ui.prpBox.setChecked(False)
        self.ui.cqcNumEdit.setFocus()

    def busy(self):
        self.ui.transferButton.setEnabled(False)
        self.ui.clearButton.setEnabled(False)
        self.ui.emailButton.setEnabled(False)
        self.ui.cqcList.setEnabled(False)

    def release(self):
        self.ui.transferButton.setEnabled(True)
        self.ui.clearButton.setEnabled(True)
        self.ui.emailButton.setEnabled(True)
        self.ui.cqcList.setEnabled(True)

class transferThread(QThread):
    result_signal = pyqtSignal(str)
    def __init__(self, cs: CQCSniffer.CQCSniffer, mode: list, cqc_info: list):
        super(transferThread, self).__init__()
        self.cs = cs
        self.mode = mode
        self.cqc_info = cqc_info
        self.tst_flag = False

    def run(self):
        try:
            cqc_num, cqe, pe, part_name = self.cqc_info
            if self.cs.checkActive():
                if self.mode[0]:
                    if self.cs.closeEvent(cqc_num, 'CQPR', cqe, 'PRP', 'The CQC sample is cleaned and prepared. The event is closed by Tianjin BL Quality COTS.'):
                        self.result_signal.emit('PRP closed. ')
                    else:
                        self.result_signal.emit('PRP not closed. ')
                if self.mode[1]:
                    if self.cs.createEvent(cqc_num, 'CQPR', cqe, pe, 'TST', 'Send the CQC part to ATE test. The event is created by Tianjin BL Quality COTS.'):
                        self.result_signal.emit('TST created. ')
                        self.tst_flag = True
                    else:
                        self.result_signal.emit('TST not created.')
                self.result_signal.emit('100')
                
            else:
                self.result_signal.emit('103')
        except Exception as err:
            print(err)
            self.progress_signal.emit('101')


class emailThread(QThread):
    result_signal = pyqtSignal(str)
    def __init__(self, cs: CQCSniffer.CQCSniffer, queue, cqeTable, peTable):
        super(emailThread, self).__init__()
        self.queue = queue
        self.cqeTable = cqeTable
        self.peTable = peTable
        self.cs = cs
    def run(self):
        try:
            date = str(datetime.today().date())
            pythoncom.CoInitialize()
            obj = Dispatch('Outlook.Application')
            mail = obj.CreateItem(0)
            mail.Subject = 'CQCs Prepared for Collection '+date
            to_list = []
            cc_list = ['helen.zhu@nxp.com;ricky.li@nxp.com;wayne.li@nxp.com;shuyuan.chai@nxp.com;xuejie.zhang@nxp.com;zhang.rui@nxp.com;yan.mu@nxp.com;da.sun@nxp.com','z.wang@nxp.com','van.fan@nxp.com','zhi.zhao@nxp.com']
            for i, row in self.queue.iterrows():
                if row['PE'] in self.peTable['PE_NAME'].values:
                    email = self.peTable[self.peTable['PE_NAME']==row['PE']]['PE_EMAIL'].iloc[0]
                    if email not in to_list:
                        to_list.append(email)
                    email = self.peTable[self.peTable['PE_NAME']==row['PE']]['MANAGER_EMAIL'].iloc[0]
                    if email not in to_list:
                        to_list.append(email)
                else:
                    email = self.cs.getEmail(row['PE'])
                    if email != None and (email not in to_list):
                        to_list.append(email)
                if row['CQE'] in self.cqeTable['CQE_NAME'].values:
                    email = self.cqeTable[self.cqeTable['CQE_NAME']==row['CQE']]['CQE_EMAIL'].iloc[0]
                    if email not in cc_list:
                        cc_list.append(email)
                else:
                    email = self.cs.getEmail(row['CQE'])
                    if email != None and (email not in cc_list):
                        cc_list.append(email)
            mail.To = ';'.join(to_list)
            mail.CC = ';'.join(cc_list)
            def to_html(table: pd.DataFrame):
                t = HTMLTable()
                l = list()
                t.append_header_rows([table.columns.values.tolist()])
                for index, row in table.iterrows():
                    l.append(row.to_list())  
                t.append_data_rows(l)
                t.set_cell_style({
                            'border-color': '#000',
                            'border-width': '1px',
                            'border-style': 'solid',
                            'border-collapse': 'collapse',
                            'padding':'4'
                        })
                t.set_header_row_style({'background-color': '#7bb1db'})
                return t.to_html()
            mail.HTMLBody = '<p>Dear Team,</p><p>Please collect your CQCs at the reception center (temporary working area on the ground floor of ATTJ).<p>&nbsp;</p>' + to_html(self.queue) + '<p>&nbsp;</p><p>&nbsp;</p><p>If you are not the responsible contact for the product, please contact Van Fan for correction.</p><p>&nbsp;</p><p>Best Regards,</p><p>Tianjin Business Line Quality</p><p>CQC Operation Tracking System</p>'
            mail.Save()
            self.result_signal.emit('100')
        except Exception as err:
            print(err)
            self.result_signal.emit('101')




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
