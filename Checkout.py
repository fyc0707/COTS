import os
from datetime import datetime

import pandas as pd
from PyQt5.QtCore import Qt, pyqtSignal, pyqtSlot
from PyQt5.QtWidgets import QCompleter, QDialog, QErrorMessage

import Ui_checkout


class Checkout(QDialog):
    def __init__(self):
        super().__init__()
        self.ui = Ui_checkout.Ui_Dialog()
        self.ui.setupUi(self)
        self.em = QErrorMessage(self)
        self.em.setWindowTitle('Error')
        self.log_file = 'log/'+datetime.today().date().isoformat()+'/log.csv'
        self.checkFile()
        self.ui.destEdit.addItems(['PE', 'FA Lab', 'Ship out', 'Others'])
        self.ui.cqcNumEdit.setFocus()
        self.data = None        

    def checkOut(self):
        if self.ui.cqcNumEdit.text() == '':
            self.em.showMessage('Please specify CQC number.')
            return
        if self.ui.destEdit.currentText() == '':
            self.em.showMessage('Please specify destination.')
            return
        try:
            checkout_time = datetime.now()
            if self.data == None:
                cqc_num = self.ui.cqcNumEdit.text()
                pe = self.ui.peEdit.text()
                if pe in self.engTable['NAME'].values:
                    pem = self.engTable[self.engTable['NAME']==pe]['MANAGER'].iloc[0]
                else:
                    pem = ''
                part_name = self.ui.partNameEdit.text()
                cqe = self.ui.cqeEdit.text()
                dest = self.ui.destEdit.currentText()
                row = []
                if cqc_num in self.df['CQC#'].values:
                    temp = self.df[self.df['CQC#']==cqc_num]
                    temp = temp[temp['Checkout']=='']
                    if not len(temp) == 0:
                        index = temp.index.to_list()[-1]
                        for col, x in self.df.iloc[index].iteritems():
                            if x=='':
                                if col == 'CQE':
                                    x = cqe
                                elif col == 'PE':
                                    x = pe
                                elif col == 'Product':
                                    x = part_name
                                elif col == 'PE Manager':
                                    x = pem
                                elif col == 'Checkout':
                                    x = 'Y'
                                elif col == 'Checkout Time':
                                    x = checkout_time.strftime('%d/%m/%Y %H:%M')
                                elif col == 'Destination':
                                    x = dest
                            else:
                                    if col == 'Status':
                                        x = 'R'                                
                            row.append(x)
                        self.df.iloc[index] = row
                    else:
                        self.df.loc[len(self.df)] = [cqc_num, '', cqe, pe, pem, '', part_name, '', '','', '', '','', 'Y', '', checkout_time.strftime('%d/%m/%Y %H:%M'), dest]
                else:
                    self.df.loc[len(self.df)] = [cqc_num, '', cqe, pe, pem, '', part_name, '', '', '', '', '','', 'Y', '', checkout_time.strftime('%d/%m/%Y %H:%M'), dest]
            else:
                cqc_num, qty, code, ship, cqe, pe, pem, part_name, ins, rcv, prp, time = self.data
                time = datetime.fromtimestamp(float(time)).strftime('%d/%m/%Y %H:%M')
                dest = self.ui.destEdit.currentText()
                row = []
                self.data = None
                if cqc_num in self.df['CQC#'].values:
                    temp = self.df[self.df['CQC#']==cqc_num]
                    temp = temp[temp['Checkout']=='']
                    if not len(temp) == 0:
                        index = temp.index.to_list()[-1]
                        for col, x in self.df.iloc[index].iteritems():
                            if x=='':
                                if col == 'CQE':
                                    x = cqe
                                elif col == 'qty':
                                    x = qty
                                elif col == 'PE':
                                    x = pe
                                elif col == 'Product':
                                    x = part_name
                                elif col == 'PE Manager':
                                    x = pem
                                elif col == 'Instruction':
                                    x = ins
                                elif col == 'Trace Code':
                                    x = code
                                elif col == 'Ship Ref.':
                                    x = ship
                                elif col == 'RCV':
                                    x = rcv
                                elif col == 'PRP':
                                    x = prp
                                elif col == 'Checkin':
                                    x = 'N'
                                elif col == 'Status':
                                    x = 'R'
                                elif col == 'Checkout':
                                    x = 'Y'
                                elif col == 'Checkin Time':
                                    x = time
                                elif col == 'Checkout Time':
                                    x = checkout_time.strftime('%d/%m/%Y %H:%M')
                                elif col == 'Destination':
                                    x = dest
                            else:
                                    if col == 'Status':
                                        x = 'R'                               
                            row.append(x)
                        self.df.iloc[index] = row
                    else:
                        self.df.loc[len(self.df)] = [cqc_num, qty, cqe, pe, pem, ins, part_name, code, ship, rcv, prp, 'Y', 'R', 'Y', time, checkout_time.strftime('%d/%m/%Y %H:%M'), dest]
                else:
                    self.df.loc[len(self.df)] = [cqc_num, qty, cqe, pe, pem, ins, part_name, code, ship, rcv, prp, 'N', 'R', 'Y', time, checkout_time.strftime('%d/%m/%Y %H:%M'), dest]
            self.df.to_csv(self.log_file, index_label=False, index=False)
            self.reset()
            self.ui.resultLabel.setText(cqc_num+' checked out.')
        except Exception as err:
            print(err)
            self.reset()
            self.ui.resultLabel.setText(cqc_num+' checkout failed.')

    @pyqtSlot()
    def on_clearButton_clicked(self):
        self.reset()
        self.data = None

    @pyqtSlot()
    def on_listenerButton_clicked(self):
        self.data = self.ui.cqcNumEdit.text()
        self.reset()
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
                    self.ui.checkOutButton.setFocus()
                else:
                    self.data = None
                    self.em.showMessage('Unidentified QR code.')
        except Exception as err:
            self.data = None
            print(err)

    def checkFile(self):
        try:
            if not os.path.exists(self.log_file):
                df = pd.DataFrame(columns=['CQC#','Qty','CQE','PE','PE Manager','Instruction','Product','Trace Code','Ship Ref.','RCV','PRP','Checkin','Status','Checkout','Checkin Time','Checkout Time','Destination'])
                df.to_csv(self.log_file, index_label=False, index=False)
            else:
                df = pd.read_csv(self.log_file, keep_default_na=False)
                if len(df) == 0:
                    df = pd.DataFrame(columns=['CQC#','Qty','CQE','PE','PE Manager','Instruction','Product','Trace Code','Ship Ref.','RCV','PRP','Checkin','Status','Checkout','Checkin Time','Checkout Time','Destination'])
                    df.to_csv(self.log_file, index_label=False, index=False)
            self.df = pd.read_csv(self.log_file, keep_default_na=False)
        except Exception as err:
            self.em.showMessage('The log file is being used by another process. Please close the file and retry. Please also delete the log.csv file if empty.')
            print(err)
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
            self.engTable = pd.read_csv('tables/EmployeeTable.csv', keep_default_na=False)
            completer = QCompleter(self.engTable[self.engTable['FUNCTION'].str.contains('PE|TECHNICIAN')]['NAME'].values.tolist())
            completer.setFilterMode(Qt.MatchContains)
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            self.ui.peEdit.setCompleter(completer)
            completer = QCompleter(self.engTable[self.engTable['FUNCTION']=='CQE']['NAME'].values.tolist())
            completer.setFilterMode(Qt.MatchContains)
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            self.ui.cqeEdit.setCompleter(completer)
        except Exception as err:
            self.em.showMessage('Failed to load the employee table. Please close the file in use and restart the window.')
            print(err)

    def destSelected(self):
        index = self.ui.destEdit.currentIndex()
        if index == 3:
            self.ui.destEdit.setEditable(True)
            self.ui.destEdit.setEditText('')
        else:
            self.ui.destEdit.setEditable(False)

    def reset(self):
        '''Reset the panel
        '''
        self.ui.cqcNumEdit.clear()
        self.ui.partNameEdit.clear()
        self.ui.cqeEdit.clear()
        self.ui.peEdit.clear()
        self.ui.destEdit.setCurrentIndex(0)
        self.ui.destEdit.setEditable(False)
        self.ui.cqcNumEdit.setFocus()
        self.ui.resultLabel.setText('')
