import os
from datetime import datetime

import pandas as pd
import pythoncom
pythoncom.CoInitialize()
from PyQt5.QtCore import QAbstractTableModel, Qt, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtWidgets import QDialog, QErrorMessage, QHeaderView, QMessageBox
from win32com.client import Dispatch
from HTMLTable import HTMLTable

import CQCSniffer
import Ui_report


class Report(QDialog):
    def __init__(self, cs: CQCSniffer.CQCSniffer):
        super().__init__()
        self.ui = Ui_report.Ui_Dialog()
        self.ui.setupUi(self)
        self.cs = cs
        self.em = QErrorMessage(self)
        self.em.setWindowTitle('Error')
        self.ui.welcomeLabel.setText('Welcome, ' + self.cs.user_name)
        self.list_file = 'log/'+datetime.today().date().isoformat()+'/wipList.xlsx'
        self.log_file = 'log/'+datetime.today().date().isoformat()+'/log.csv'
        self.checkFile()

    def openWith(self):
        try:
            os.startfile(os.path.abspath(self.log_file))
        except Exception as err:
            self.em.showMessage('Failed to open the log file. Please close the file used by other processes.')
            print(err)

    def checkFile(self):
        try:
            if os.path.exists(self.log_file):
                df = pd.read_csv(self.log_file, keep_default_na=False)
                if len(df) == 0:
                    df = pd.DataFrame(columns=['CQC#','Qty','CQE','PE','PE Manager','Instruction','Product','Trace Code','Ship Ref.','RCV','PRP','Checkin','Status','Checkout','Checkin Time','Checkout Time','Destination'])
                    df.to_csv(self.log_file, index_label=False, index=False)
            else:
                df = pd.DataFrame(columns=['CQC#','Qty','CQE','PE','PE Manager','Instruction','Product','Trace Code','Ship Ref.','RCV','PRP','Checkin','Status','Checkout','Checkin Time','Checkout Time','Destination'])
                df.to_csv(self.log_file, index_label=False, index=False)
            self.df = pd.read_csv(self.log_file, keep_default_na=False)
            self.updateTable()
            
        except Exception as err:
            self.em.showMessage('The log file is being used by another process. Please close the file and retry. Please also delete the log.csv file if empty.')
            print(err)

        try:
            self.productTable = pd.read_csv('tables/ProductTable.csv', keep_default_na=False)
        except Exception as err:
            self.em.showMessage('Failed to load the product table. Please close the file in use and restart the window.')
            print(err)
        try:
            self.engTable = pd.read_csv('tables/EmployeeTable.csv', keep_default_na=False)
        except Exception as err:
            self.em.showMessage('Failed to load the employee table. Please close the file in use and restart the window.')
            print(err)
        
    def email(self):
        if len(self.df)==0:
            self.em.showMessage('The list is empty.')
            return
        self.thread = emailThread(self.cs, self.df, self.engTable, self.productTable)
        self.thread.result_signal.connect(self.emailCallBack)
        self.thread.start()
        self.busy()
    
    def emailCallBack(self, signal):
        if signal == '100':
            self.ui.resultLabel.setText('Email sent successfully.')
            self.release()
        else:
            self.ui.resultLabel.setText('Email failed.')
            self.release()

    def valueChanged(self, row, col, value):
        try:
            self.df.iloc[row,col] = value
            self.df.to_csv(self.log_file, index_label=False, index=False)
        except Exception as err:
            print(err)
            self.em.showMessage('The log file is being used by another process. Please close the file and retry.')
            self.updateTable()

    def updateTable(self):
        self.model = pandasModel(self.df)
        self.model.value_signal.connect(self.valueChanged)
        self.ui.cqcList.setModel(self.model)
        self.ui.cqcList.horizontalHeader().setDefaultAlignment(Qt.AlignCenter | Qt.Alignment(Qt.TextWordWrap))
        self.ui.cqcList.horizontalHeader().setFixedHeight(40)
        self.ui.cqcList.setColumnWidth(0,60)
        self.ui.cqcList.setColumnWidth(1,20)
        self.ui.cqcList.setColumnWidth(2,100)
        self.ui.cqcList.setColumnWidth(3,100)
        self.ui.cqcList.setColumnWidth(4,100)
        self.ui.cqcList.setColumnWidth(5,100)
        self.ui.cqcList.setColumnWidth(6,110)
        self.ui.cqcList.setColumnWidth(7,80)
        self.ui.cqcList.setColumnWidth(8,110)
        self.ui.cqcList.setColumnWidth(9,30)
        self.ui.cqcList.setColumnWidth(10,30)
        self.ui.cqcList.setColumnWidth(11,60)
        self.ui.cqcList.setColumnWidth(12,60)
        self.ui.cqcList.setColumnWidth(13,60)
        self.ui.cqcList.setColumnWidth(14,80)
        self.ui.cqcList.setColumnWidth(15,80)
        self.ui.cqcList.setColumnWidth(16,80)

    def busy(self):
        self.ui.emailButton.setEnabled(False)
        self.ui.openButton.setEnabled(False)
        self.ui.refreshButton.setEnabled(False)

    def release(self):
        self.ui.emailButton.setEnabled(True)
        self.ui.openButton.setEnabled(True)
        self.ui.refreshButton.setEnabled(True)

class emailThread(QThread):
    result_signal = pyqtSignal(str)
    def __init__(self, cs: CQCSniffer.CQCSniffer, df, engTable, productTable):
        super(emailThread, self).__init__()
        self.df = df
        self.engTable = engTable
        self.productTable = productTable
        self.cs = cs

    def run(self):
        try:
            pythoncom.CoInitialize()
            date = str(datetime.today().date())
            obj = Dispatch('Outlook.Application')
            mail = obj.CreateItem(0)
            mail.Subject = 'Report of Received CQCs '+date
            to_list = []
            cc_list = []
            cc_list.extend(self.engTable[self.engTable['GM_RCV']=='Y']['EMAIL'].to_list())
            for i, row in self.df.iterrows():
                if row['Product'] in self.productTable['PART_TYPE_NAME'].values:
                    r = self.productTable[self.productTable['PART_TYPE_NAME']==row['Product']].iloc[0]
                    atts = r['ATTENTION_NAME']
                    for att in atts.split(';'):
                        if att != '':
                            email = self.engTable[self.engTable['NAME']==att]['EMAIL'].iloc[0]
                            if email not in cc_list:
                                cc_list.append(email)
                if row['PE'] in self.engTable['NAME'].values:
                    r = self.engTable[self.engTable['NAME']==row['PE']].iloc[0]
                    email = r['EMAIL']
                    if email not in to_list:
                        to_list.append(email)
                    email = r['MANAGER_EMAIL']
                    if email not in cc_list:
                        cc_list.append(email)
                    atts = r['ATTENTION_NAME']
                    for att in atts.split(';'):
                        if att != '':
                            email = self.engTable[self.engTable['NAME']==att]['EMAIL'].iloc[0]
                            if email not in to_list:
                                to_list.append(email)
                elif row['PE']!='':
                    name, email, mgr, mgr_email = self.cs.getFullInfo(row['PE'])
                    if email != '' and (email not in to_list):
                        to_list.append(email)
                    if mgr_email != '' and (mgr_email not in cc_list):
                        cc_list.append(mgr_email)
                    if name != '' and email != '' and mgr != '' and mgr_email != '':
                        try:
                            self.engTable.loc[len(self.engTable)] = [name, email, 'PE', mgr, mgr_email, '', '', '']
                            self.engTable.to_csv('tables/EmployeeTable.csv', index_label=False, index=False)
                        except Exception as err:
                            print(err)
                if row['CQE'] in self.engTable['NAME'].values:
                    r = self.engTable[self.engTable['NAME']==row['CQE']].iloc[0]
                    email = r['EMAIL']
                    if email not in to_list:
                        to_list.append(email)
                    email = r['MANAGER_EMAIL']
                    if email not in cc_list:
                        cc_list.append(email)
                    atts = r['ATTENTION_NAME']
                    for att in atts.split(';'):
                        if att != '':
                            email = self.engTable[self.engTable['NAME']==att]['EMAIL'].iloc[0]
                            if email not in to_list:
                                to_list.append(email)
                elif row['CQE']!='':
                    name, email, mgr, mgr_email = self.cs.getFullInfo(row['CQE'])
                    if email != '' and (email not in to_list):
                        to_list.append(email)
                    if mgr_email != '' and (mgr_email not in cc_list):
                        cc_list.append(mgr_email)
                    if name != '' and email != '' and mgr != '' and mgr_email != '':
                        try:
                            self.engTable.loc[len(self.engTable)] = [name, email, 'CQE', mgr, mgr_email, '', '', '']
                            self.engTable.to_csv('tables/EmployeeTable.csv', index_label=False, index=False)
                        except Exception as err:
                            print(err)
            mail.To = ';'.join(to_list)
            mail.CC = ';'.join(cc_list)    
            self.df = self.df.astype(str)
            table_cleared = self.df[self.df['Checkout']=='Y']
            table_underway = self.df[(self.df['Status']=='P') & (self.df['Checkout']=='')]
            table_ready = self.df[(self.df['Status']=='R') & (self.df['Checkout']=='')]
            table_store = self.df[(self.df['Status']=='S') & (self.df['Checkout']=='')]
            bg = ['#f9b500','#7bb1db','#c9d200','#00a4a7']
            output = [table_ready, table_underway, table_cleared, table_store]
            for i in range(4):
                t = HTMLTable()
                l = list()
                t.append_header_rows([output[i].columns.values.tolist()])
                if len(output[i]) == 0:
                    output[i] = None
                else:   
                    for index, row in output[i].iterrows():
                        l.append(row.to_list())
                    t.append_data_rows(l)
                    t.set_cell_style({
                                'border-color': '#aaaaaa',
                                'border-width': '1px',
                                'border-style': 'solid',
                                'border-collapse': 'collapse',
                                'padding':'4'
                            })
                    t.set_header_row_style({'background-color': bg[i]})
                    output[i] = t.to_html()
            mail.HTMLBody = '<p>Dear Team,</p><p>Please refer to the list(s) of the '+str(len(self.df))+' CQC(s) that have been handled at the reception center today. For the un-checkout CQCs, please arrange resources for sample preparation and verification according to the instruction. For the CQCs that need sample cleaning, notification emails will be sent to the responsible engineers when the CQCs are ready to collect.<p>&nbsp;</p>' + (('<p style="color:black">' + str(len(table_ready)) + ' CQC(s) are waiting to be collected.</p>' + output[0]) if output[0] != None else '<p style="color:black">No CQC is waiting to be collected.</p>') + (('<p style="color:black">' + str(len(table_underway)) + ' CQC(s) are under preparation.</p>' + output[1]) if output[1] != None else '<p style="color:black">No CQC is under preparation.</p>') + (('<p style="color:black">' + str(len(table_cleared)) + ' CQC(s) have been checked out.</p>' + output[2]) if output[2] != None else '<p style="color:black">No CQC was checked out.</p>') + (('<p style="color:black">' + str(len(table_store)) + ' CQC(s) have been stored.</p>' + output[3]) if output[3] != None else '<p style="color:black">No CQC was stored.</p>') + '<p>&nbsp;</p><p>&nbsp;</p><p>If you are not the responsible contact for the product, please contact Van Fan for correction.</p><p>&nbsp;</p><p>Best Regards,</p><p>Tianjin Business Line Quality</p><p>CQC Operation Tracking System</p>'
            mail.Save()
            self.result_signal.emit('100')
        except Exception as err:
            print(err)
            self.result_signal.emit('101')


class pandasModel(QAbstractTableModel):
    '''Model dataframe to QTableView
    '''
    value_signal = pyqtSignal(int, int, str)
    def __init__(self, data):
        super(pandasModel, self).__init__()
        self.data = data
    
    def rowCount(self, parent=None):
        return self.data.shape[0]
    
    def columnCount(self, parent=None):
        return self.data.shape[1]
    
    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole or role == Qt.EditRole:
                return str(self.data.iloc[index.row(), index.column()])
        return None
    
    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.data.columns[section]
        return None
    
    def setData(self, index, value, role):
        if not index.isValid():
            return False
        if role != Qt.EditRole:
            return False
        row = index.row()
        if row < 0 or row >= len(self.data.values):
            return False
        column = index.column()
        if column < 0 or column >= self.data.columns.size:
            return False
        self.data.iloc[row,column] = value
        self.dataChanged.emit(index, index)
        self.value_signal.emit(int(row), int(column), str(value))
        return True
    
    def flags(self, index):
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable