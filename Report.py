import os
from datetime import datetime

import pandas as pd
import pythoncom
from PyQt5.QtCore import QAbstractTableModel, Qt, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtWidgets import QDialog, QErrorMessage, QHeaderView, QMessageBox
from win32com.client import Dispatch

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
        pythoncom.CoInitialize()
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
                    df = pd.DataFrame(columns=['CQC#','Qty','CQE','PE','PE Manager','Instruction','Product','Trace Code','Ship Ref.','RCV','PRP','Checkin','Checkout','Checkin Time','Checkout Time','Destination'])
                    df.to_csv(self.log_file, index_label=False, index=False)
            else:
                df = pd.DataFrame(columns=['CQC#','Qty','CQE','PE','PE Manager','Instruction','Product','Trace Code','Ship Ref.','RCV','PRP','Checkin','Checkout','Checkin Time','Checkout Time','Destination'])
                df.to_csv(self.log_file, index_label=False, index=False)
            self.df = pd.read_csv(self.log_file, keep_default_na=False)
            self.updateTable()
            
        except Exception as err:
            self.em.showMessage('Failed to load the log file. Please close the file used by other processes.')
            print(err)

        try:
            self.productTable = pd.read_csv('ProductTable.csv', keep_default_na=False)
        except Exception as err:
            self.em.showMessage('Failed to load the product table. Please close the file in use and restart the window.')
            print(err)
        try:
            self.peTable = pd.read_csv('PETable.csv', keep_default_na=False)
        except Exception as err:
            self.em.showMessage('Failed to load the PE table. Please close the file in use and restart the window.')
            print(err)
        try:
            self.cqeTable = pd.read_csv('CQETable.csv', keep_default_na=False)
        except Exception as err:
            self.em.showMessage('Failed to load the CQE table. Please close the file in use and restart the window.')
            print(err)

    def email(self):
        if len(self.df)==0:
            self.em.showMessage('The list is empty.')
            return
        self.thread = emailThread(self.cs, self.df, self.cqeTable, self.peTable)
        self.thread.result_signal.connect(self.emailCallBack)
        pythoncom.CoInitialize()
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


    def updateTable(self):
        model = pandasModel(self.df)
        self.ui.cqcList.setModel(model)
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
        self.ui.cqcList.setColumnWidth(13,80)
        self.ui.cqcList.setColumnWidth(14,80)
        self.ui.cqcList.setColumnWidth(15,80)

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
    def __init__(self, cs: CQCSniffer.CQCSniffer, df, cqeTable, peTable):
        super(emailThread, self).__init__()
        self.df = df
        self.cqeTable = cqeTable
        self.peTable = peTable
        self.cs = cs
        pythoncom.CoInitialize()
    def run(self):
        try:
            date = str(datetime.today().date())
            obj = Dispatch('Outlook.Application')
            mail = obj.CreateItem(0)
            mail.Subject = 'Report of Received CQCs '+date
            to_list = []
            cc_list = ['helen.zhu@nxp.com;ricky.li@nxp.com;wayne.li@nxp.com;shuyuan.chai@nxp.com;xuejie.zhang@nxp.com;zhang.rui@nxp.com;yan.mu@nxp.com;da.sun@nxp.com','z.wang@nxp.com','van.fan@nxp.com','zhi.zhao@nxp.com']
            for i, row in self.df.iterrows():
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
                    if email not in to_list:
                        to_list.append(email)
                else:
                    email = self.cs.getEmail(row['CQE'])
                    if email != None and (email not in to_list):
                        to_list.append(email)
            mail.To = ';'.join(to_list)
            mail.CC = ';'.join(cc_list)
            mail.HTMLBody = '<p>Dear Team,</p><p>Please refer to the CQCs that were received at the reception center today. For the un-checkout CQCs, please arrange resources for sample preparation and verification according to the instruction.<p>&nbsp;</p>' + self.df.to_html(escape=False, na_rep='N/A', border=1) + '<p>&nbsp;</p><p>&nbsp;</p><p>If you are not the responsible contact for the product, please contact Van Fan for correction.</p><p>&nbsp;</p><p>Best Regards,</p><p>Tianjin Business Line Quality</p><p>CQC Operation Tracking System</p>'
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
