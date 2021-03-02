#coding=utf-8
import json
import os
import re
from datetime import datetime

import pandas as pd
import pythoncom
import requests
from dateutil import parser

pythoncom.CoInitialize()
from PyQt5.QtCore import QAbstractTableModel, Qt, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtWidgets import QDialog, QErrorMessage
from win32com.client import Dispatch
from HTMLTable import HTMLTable

import CQCSniffer
import Ui_shipment


class Shipment(QDialog):
    def __init__(self, cs: CQCSniffer.CQCSniffer):
        super().__init__()
        self.ui = Ui_shipment.Ui_Dialog()
        self.ui.setupUi(self)
        self.em = QErrorMessage(self)
        self.em.setWindowTitle('Error')
        self.cs = cs
        self.ui.welcomeLabel.setText('Welcome, ' + self.cs.user_name)
        self.list_file = 'log/'+datetime.today().date().isoformat()+'/wipList.xlsx'
        self.ship_file = 'log/'+datetime.today().date().isoformat()+'/shipment.csv'
        self.checkFile()

    def checkFile(self):
        try:
            self.productTable = pd.read_csv('tables/ProductTable.csv', keep_default_na=False)
        except Exception as err:
            self.em.showMessage('Failed to load the product table. Please close the file in use and restart the window.')
            print(err)
        try:
            self.engTable = pd.read_csv('tables/EmployeeTable.csv', keep_default_na=False)
        except Exception as err:
            self.em.showMessage('Failed to load the engineer table. Please close the file in use and restart the window.')
            print(err)
        try: 
            if os.path.exists(self.ship_file):
                self.df = pd.read_csv(self.ship_file, keep_default_na=False)
                self.ui.cqcListLable.setText('Last Update:\n'+datetime.fromtimestamp(os.path.getmtime(self.ship_file)).strftime('%Y-%m-%d %H:%M:%S'))
                model = pandasModel(self.df)
                model.value_signal.connect(self.valueChanged)
                self.ui.cqcList.setModel(model)
                self.ui.cqcList.setColumnWidth(0,60)
                self.ui.cqcList.setColumnWidth(1,100)
                self.ui.cqcList.setColumnWidth(2,90)
                self.ui.cqcList.setColumnWidth(3,90)
                self.ui.cqcList.setColumnWidth(4,90)
                self.ui.cqcList.setColumnWidth(5,20)
                self.ui.cqcList.setColumnWidth(6,90)
                self.ui.cqcList.setColumnWidth(7,150)
                self.ui.cqcList.setColumnWidth(8,60)
                self.ui.cqcList.setColumnWidth(9,90)
                self.ui.cqcList.setColumnWidth(10,100)
                self.ui.cqcList.setColumnWidth(11,100)
                self.ui.cqcList.setColumnWidth(12,90)
                self.ui.cqcList.setColumnWidth(13,110)
                self.ui.cqcList.setColumnWidth(14,100)            
            else:
                self.ui.cqcListLable.setText('No list found')
                self.df = None
        except Exception as err:
            print(err)
            self.em.showMessage('Failed to load the shipment table. Please close the file in use and restart the window.')
    
    def itemSelected(self):
        self.ui.shipperLink.clear()
        row = self.ui.cqcList.selectedIndexes()[0].row()
        carrier = self.df['Carrier'].loc[row]
        num = self.df['Ship Ref.'].loc[row]
        if carrier != '' and num != '':
            if carrier == 'FedEx CN':
                self.ui.shipperLink.setText('<a href="https://cndxp.apac.fedex.com/app/track?method=query&language=en&region=CN&tn=%s">Track %s on FedEx CN</a>'%(num,num))
            if carrier == 'TNT':
                self.ui.shipperLink.setText('<a href="https://www.tnt.com/express/en_cn/site/shipping-tools/tracking.html?searchType=CON&cons=%s">Track %s on TNT</a>'%(num,num))
            if carrier == 'SF':
                self.ui.shipperLink.setText('<a href="https://www.sf-express.com/cn/en/dynamic_function/waybill/#search/bill-number/%s">Track %s on SF</a>'%(num,num))
            if carrier == 'UPS':
                self.ui.shipperLink.setText('<a href="https://www.ups.com/track?loc=en_CN&tracknum=%s">Track %s on UPS</a>'%(num,num))
            if carrier == 'EMS':
                self.ui.shipperLink.setText('<a href="http://www.ems.com.cn/mailtracking/e_you_jian_cha_xun.html">Track %s on EMS</a>'%(num))

    def updateCQCList(self):
        self.ui.progressBar.setValue(0)
        self.thread = downloadThread(self.cs, self.list_file, self.ship_file, self.productTable, 1)
        self.thread.process_signal.connect(self.downloadCallBack)
        self.thread.start()
        self.ui.cqcListLable.setText('Downloading')
        self.busy()

    def downloadCallBack(self, signal):
        if signal == 101:
            self.release()
            self.checkFile()
            self.em.showMessage('Download failed. Please retry or restart the application.')
            self.ui.cqcListLable.setText('No list found')
            self.ui.progressBar.setValue(0)
            os.remove(self.list_file+'1')
        elif signal == 104:
            self.release()
            self.checkFile()
            self.em.showMessage('Download failed. Please contact COTS admin.')
            self.ui.cqcListLable.setText('No list found')
            self.ui.progressBar.setValue(0)
        elif signal == 102:
            self.release()
            try:
                os.replace(self.list_file+'1', self.list_file)
                os.replace(self.ship_file+'1', self.ship_file)
            except:
                self.em.showMessage('The list file is being used by another process. Please close the file and retry.')
                self.ui.cqcListLable.setText('No list found')
            self.checkFile()
        else:
            self.ui.progressBar.setValue(signal)

    def sendEmail(self):
        self.ui.progressBar.setValue(0)
        if len(self.df)==0:
            self.em.showMessage('The list is empty.')
            return
        self.thread = emailThread(self.cs, self.df, self.engTable)
        self.thread.result_signal.connect(self.emailCallBack)
        self.thread.start()
        self.busy()
    
    def emailCallBack(self, signal):
        if signal == '100':
            self.ui.cqcListLable.setText('Email sent successfully.')
            self.release()
        else:
            self.ui.cqcListLable.setText('Email failed.')
            self.release()
    
    def busy(self):
        self.ui.emailButton.setEnabled(False)
        self.ui.getListButton.setEnabled(False)

    def release(self):
        self.ui.emailButton.setEnabled(True)
        self.ui.getListButton.setEnabled(True)
    
    def valueChanged(self, row, col, value):
        try:
            self.df.iloc[row,col] = value
            self.df.to_csv(self.ship_file, index_label=False, index=False)
        except Exception as err:
            print(err)
            self.em.showMessage('The shipment file is being used by another process. Please close the file and retry.')
            self.updateTable()

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


class downloadThread(QThread):
    '''RCV file download thread
    '''
    process_signal = pyqtSignal(int)

    def __init__(self, cs, list, ship, product, buffer):
        super(downloadThread, self).__init__()
        self.cs = cs
        self.filesize = None
        self.list_file = list
        self.buffer = buffer
        self.ship_file = ship
        self.productTable = product
        self.fileobj = open(self.list_file+'1', 'wb')

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
                    process = offset / int(size) * 80
                    self.process_signal.emit(int(process))
                self.process_signal.emit(80)
            else:
                self.process_signal.emit(101)
        except:
            self.process_signal.emit(101)
        self.fileobj.close()
        try:
            self.df = self.cs.getWIPData(self.list_file+'1')
            self.df = self.df[(self.df['Type']=='CQPR') & (self.df['Event'].str.contains('RCT|RCV'))]
            self.df.reset_index(drop=True, inplace=True)
            self.df = self.df[['CQC#','CQE','Customer','Part Name','Qty','Trace Code','Instruction','Ship Ref.']]
            self.df.insert(4,'PE','')
            self.df.insert(7,'Carrier','')
            for index, row in self.df.iterrows():
                if row['Ship Ref.'] != '':
                    trim = str(row['Ship Ref.']).replace(' ','').replace(':','').upper()
                    trim = re.sub(r'[^a-zA-Z0-9]','',trim)
                    if 'FED' in trim:
                        row['Carrier'] = 'FedEx'
                        trim = re.sub(r'[A-Z]','',trim)
                        if trim.startswith('12'):
                            row['Carrier'] = 'FedEx CN'
                        row['Ship Ref.'] = trim
                    elif 'DHL' in trim:
                        row['Carrier'] = 'DHL'
                        trim = re.sub(r'[A-Z]','',trim)
                        row['Ship Ref.'] = trim
                    elif 'UPS' in trim:
                        row['Carrier'] = 'UPS'
                        trim = trim.replace('UPS','')
                        row['Ship Ref.'] = trim
                    elif 'TNT' in trim:
                        row['Carrier'] = 'TNT'
                        trim = re.sub(r'[A-Z]','',trim)
                        row['Ship Ref.'] = trim
                    elif 'SF' in trim:
                        row['Carrier'] = 'SF'
                        trim = re.sub(r'[A-Z]','',trim)
                        row['Ship Ref.'] = 'SF'+trim
                    elif 'EMS' in trim:
                        row['Carrier'] = 'EMS'
                        trim = re.sub(r'[A-Z]','',trim)
                        row['Ship Ref.'] = trim
                    elif re.match(r'^[A-Z]{2,}[0-9]{4,}$',trim):
                        row['Carrier'] = re.search(r'^[A-Z]{2,}',trim).group()
                        row['Ship Ref.'] = re.sub(r'^[A-Z]{2,}','',trim)
                    elif re.match(r'^[0-9]{4,}[A-Z]{2,}$',trim):
                        row['Carrier'] = re.search(r'[A-Z]{2,}$',trim).group()
                        row['Ship Ref.'] = re.sub(r'[A-Z]{2,}$','',trim)
                    elif re.search(r'^[0-9]*$',trim):
                        row['Carrier'] = 'N/A'
                    else:
                        row['Carrier'] = 'N/A'
                if row['Part Name'] in self.productTable['PART_TYPE_NAME'].values:
                    row['PE'] = self.productTable[self.productTable['PART_TYPE_NAME']==row['Part Name']]['PE_NAME'].iloc[0]
            result_df = pd.DataFrame(columns=['Carrier','Ship Ref.','Origin', 'Ship Date', 'Destination','Status','Current Location','Delivery Date'])
            result_df[['Carrier','Ship Ref.']] = self.df.drop_duplicates(['Carrier','Ship Ref.'])[['Carrier','Ship Ref.']]

            #FedEx Tracking
            tracking_numbers = result_df[result_df['Carrier']=='FedEx']['Ship Ref.'].to_list()
            result = []
            limit = 30
            chunked = [tracking_numbers[i:i + limit] for i in range(0, len(tracking_numbers), limit)]
            for chunk in chunked:
                result.extend(self.track_fedex(chunk))
            result_df.loc[self.df['Carrier']=='FedEx',['Origin','Ship Date','Destination','Status','Current Location','Delivery Date']] = result

            #DHL Tracking
            result = []
            tracking_numbers = result_df[result_df['Carrier']=='DHL']['Ship Ref.'].to_list()
            for n in tracking_numbers:
                result.append(self.track_dhl(n))
            result_df.loc[self.df['Carrier']=='DHL',['Origin','Ship Date','Destination','Status','Current Location','Delivery Date']] = result
            self.process_signal.emit(100)
            #Others Tracking

            self.df = pd.merge(self.df, result_df, how='left', on=['Carrier', 'Ship Ref.'])
            self.df = self.df[['CQC#','CQE','Customer','Part Name','PE','Qty','Trace Code','Instruction','Carrier', 'Ship Ref.','Origin','Destination','Status','Current Location','Delivery Date']]
            self.df.to_csv(self.ship_file+'1', na_rep='',index=False)
            self.process_signal.emit(102)
        except Exception as err:
            print(err)
            self.process_signal.emit(104)

        self.exit(0)

    def track_fedex(self, tracking_numbers):
        trackingInfoList = [{'trackNumberInfo': {
            'trackingNumber': x
        }
        } for x in tracking_numbers]
        data = requests.post('https://www.fedex.com/trackingCal/track', data={
            'data': json.dumps({
                'TrackPackagesRequest': {
                'trackingInfoList': trackingInfoList
                }
            }),
            'action': 'trackpackages',
            'locale': 'en_US',
            'format': 'json',
            'version': 99
        }).json()
        r = []
        for i in range(len(tracking_numbers)):
            info   = data['TrackPackagesResponse']['packageList'][i]
            origin = (info['shipperCity'] + ', ' if info['shipperCntryCD'] != '' else '') + info['shipperCntryCD']
            dest   = (info['recipientCity'] + ', ' if info['recipientCntryCD'] != '' else '') + info['recipientCntryCD']
            shipdate = info['displayShipDt']
            try:
                if shipdate != '':
                    shipdate = datetime.strptime(shipdate,'%m/%d/%Y').strftime('%d/%m/%Y')
                else:
                    shipdate = 'Pending'
            except:
                shipdate = 'Pending'

            status = info['keyStatus']
            loc    = (info['statusLocationCity'] + ', ' if info['statusLocationCntryCD'] != '' else '') + info['statusLocationCntryCD']
            delivery_time = info["displayEstDeliveryDt"]
            if len(delivery_time)<1:
                delivery_time = info["displayActDeliveryDt"]
            if delivery_time != 'Pending':
                try:
                    delivery_time = datetime.strptime(delivery_time,'%m/%d/%Y').strftime('%d/%m/%Y')
                except:
                    pass
            else:
                delivery_time = 'Pending'
            r.append([origin, shipdate, dest, status, loc, delivery_time])
        return r

    def track_dhl(self, tracking_number):
        data = requests.get('https://www.dhl.com/shipmentTracking?AWB='+str(tracking_number),
            headers={'User-agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.96 Safari/537.36 Edg/88.0.705.49"
            }).json()
        origin, shipdate, dest, status, loc, delivery_time = '', '', '', '', '', ''
        if 'errors' in data.keys():
            return ['','','','','','']
        else:
            try:
                info = data['results'][0]
                origin = info['origin']['value']
                try:
                    shipdate = info['checkpoints'][-1]['date']
                    shipdate = parser.parse(shipdate).strftime('%d/%m/%Y')
                except:
                    pass
                dest = info['destination']['value']
                status = info['delivery']['status']
                loc = info['signature']['signatory']
                if 'edd' in info.keys():
                    try:
                        delivery_time = info['edd']['date']
                        delivery_time = parser.parse(delivery_time).strftime('%d/%m/%Y')
                    except:
                        pass
                elif status == 'delivered':
                    try:
                        delivery_time = info['signature']['description']
                        delivery_time = parser.parse(delivery_time).strftime('%d/%m/%Y')
                    except:
                        pass
            except:
                pass
            return [origin, shipdate, dest, status, loc, delivery_time]

    def track_kdn(self, tracking_number):
        pass


class emailThread(QThread):
    result_signal = pyqtSignal(str)
    def __init__(self, cs: CQCSniffer.CQCSniffer, df, engTable):
        super(emailThread, self).__init__()
        self.df = df
        self.engTable = engTable
        self.cs = cs

    def run(self):
        try:
            pythoncom.CoInitialize()
            date = str(datetime.today().date())
            obj = Dispatch('Outlook.Application')
            mail = obj.CreateItem(0)
            mail.Subject = 'CQC on the Way '+date
            to_list = []
            cc_list = ['helen.zhu@nxp.com','ricky.li@nxp.com','wayne.li@nxp.com','shuyuan.chai@nxp.com','xuejie.zhang@nxp.com','z.wang@nxp.com','van.fan@nxp.com','zhi.zhao@nxp.com']
            for i, row in self.df.iterrows():
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
                else:
                    name, email, mgr, mgr_email = self.cs.getFullInfo(row['PE'])
                    if email != '' and (email not in to_list):
                        to_list.append(email)
                    if mgr_email != '' and (mgr_email not in cc_list):
                        cc_list.append(mgr_email)
                    if name != '' and email != '' and mgr != '' and mgr_email != '':
                        try:
                            self.engTable.loc[len(self.engTable)] = [name, email, 'PE', mgr, mgr_email, '']
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
                else:
                    name, email, mgr, mgr_email = self.cs.getFullInfo(row['CQE'])
                    if email != '' and (email not in to_list):
                        to_list.append(email)
                    if mgr_email != '' and (mgr_email not in cc_list):
                        cc_list.append(mgr_email)
                    if name != '' and email != '' and mgr != '' and mgr_email != '':
                        try:
                            self.engTable.loc[len(self.engTable)] = [name, email, 'CQE', mgr, mgr_email, '']
                            self.engTable.to_csv('tables/EmployeeTable.csv', index_label=False, index=False)
                        except Exception as err:
                            print(err)
            mail.To = ';'.join(to_list)
            mail.CC = ';'.join(cc_list)    
            self.df = self.df.astype(str)
            
            def to_html(table: pd.DataFrame):
                table = table.astype(str)
                t = HTMLTable()
                l = list()
                t.append_header_rows([table.columns.values.tolist()])
                for index, row in table.iterrows():
                    l.append(row.to_list())  
                t.append_data_rows(l)
                t.set_cell_style({
                            'border-color': '#aaaaaa',
                            'border-width': '1px',
                            'border-style': 'solid',
                            'border-collapse': 'collapse',
                            'padding':'4'
                        })
                t.set_header_row_style({'background-color': '#c9d200'})
                return t.to_html()
            mail.HTMLBody = '<p>Dear Team,</p><p>Please refer to the list of the '+str(len(self.df))+' CQC(s) that are on the way to Tianjin CQC reception center, please arrange resources for sample preparation and verification.</p>'+to_html(self.df)+'<p>&nbsp;</p><p>If you are not the responsible contact for the product, please contact Van Fan for correction.</p><p>&nbsp;</p><p>Best Regards,</p><p>Tianjin Business Line Quality</p><p>CQC Operation Tracking System</p>'
            mail.Save()
            self.result_signal.emit('100')
        except Exception as err:
            print(err)
            self.result_signal.emit('101')
