#coding=utf-8
import os, subprocess, sys
from datetime import datetime
from io import BytesIO

import pandas as pd
from barcode import Code39
from barcode.writer import ImageWriter
from PyQt5.QtCore import QAbstractTableModel, Qt, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtWidgets import QDialog, QErrorMessage, QMessageBox
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Image
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle as PS

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
        self.list_file = 'log/'+datetime.today().date().isoformat()+'/wipList.xlsx'
        self.checkFile()
        
    def fillInfo(self):
        pass

    def itemSelected(self):
        row = self.ui.cqcList.selectedIndexes()[0].row()
        self.ui.cqcNumEdit.setText(self.rcv_df['CQC#'].loc[row])
        self.ui.partNameEdit.setText(self.rcv_df['Part Name'].loc[row])
        self.ui.cqeEdit.setText(self.rcv_df['CQE'].loc[row])
        self.ui.qtyEdit.setText(self.rcv_df['Qty'].loc[row])
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
            try:
                os.replace(self.list_file+'1', self.list_file)
                self.ui.resultLabel.setText('Download successfully')
            except:
                self.em.showMessage('The original file is being used by another process. Please close the file and retry.')
            self.checkFile()
        else:
            self.ui.progressBar.setValue(signal)

    def checkin(self):
        event = ''
        cqc_type = ''
        try:
            event = 'RCT' if 'RCT' in self.rcv_df[self.rcv_df['CQC#']==self.ui.cqcNumEdit.text()]['Event'].iloc[0] else 'RCV'
            cqc_type = self.rcv_df[self.rcv_df['CQC#']==self.ui.cqcNumEdit.text()]['Type'].iloc[0]
        except:
            event = ''
            cqc_type = ''
            self.ui.rcvBox.setChecked(False)
            self.ui.prpBox.setChecked(False)
        
        mode = [self.ui.checkOnlyBox.isChecked(), self.ui.printBox.isChecked(), self.ui.prpBox.isChecked(), self.ui.rcvBox.isChecked()]
        if mode == [False]*4:
            self.em.showMessage('Please check options.')
        else:           
            event = ''
            cqc_type = ''
            try:
                event = 'RCT' if 'RCT' in self.rcv_df[self.rcv_df['CQC#']==self.ui.cqeEdit.text()]['Event'].iloc[0] else 'RCV'
                cqc_type = self.rcv_df[self.rcv_df['CQC#']==self.ui.cqcNumEdit.text()]['Type'].iloc[0]
            except:
                event = ''
                cqc_type = ''
            cqc_info = [self.ui.cqcNumEdit.text(), self.ui.partNameEdit.text(), self.ui.qtyEdit.text(), 
                self.ui.cqeEdit.text(), self.ui.peEdit.text(), self.ui.instruEdit.text(), event, cqc_type]
            
            self.thread = checkinThread(self.cs, mode, cqc_info)
            self.thread.progress_signal.connect(self.checkinCallBack)
            self.thread.status_signal.connect(self.checkinCallBack)
            self.thread.start()
            self.busy()
    
    def checkinCallBack(self, signal):
        pass


    def checkFile(self):
        if os.path.exists(self.list_file):
            self.ui.cqcListLable.setText('Last Update:\n'+datetime.fromtimestamp(os.path.getmtime(self.list_file)).strftime('%Y-%m-%d %H:%M:%S'))
            self.wip_df = self.cs.getWIPData(self.list_file)
            self.rcv_df = self.wip_df[self.wip_df['Event'].str.contains('RCT|RCV')]
            self.rcv_df.reset_index(drop=True, inplace=True)
            model = pandasModel(self.rcv_df.drop(['Event', 'B2B'], axis=1))
            self.ui.cqcList.setModel(model)
            self.ui.cqcList.setColumnWidth(0,55)
            self.ui.cqcList.setColumnWidth(1,30)
            self.ui.cqcList.setColumnWidth(2,90)
            self.ui.cqcList.setColumnWidth(3,85)
            self.ui.cqcList.setColumnWidth(4,85)
            self.ui.cqcList.setColumnWidth(5,28)
            self.ui.cqcList.setColumnWidth(6,90)
            self.ui.cqcList.setColumnWidth(7,160)

        else:
            self.ui.cqcListLable.setText('No list found')
    

    def reset(self):
        '''Reset the panel
        '''
        self.ui.cqcNumEdit.clear()
        self.ui.partNameEdit.clear()
        self.ui.cqeEdit.clear()
        self.ui.peEdit.clear()
        self.ui.qtyEdit.clear()
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
        self.ui.fillInfoButton.setEnabled(False)

    def release(self):
        self.ui.getListButton.setEnabled(True)
        self.ui.checkinButton.setEnabled(True)
        self.ui.fillInfoButton.setEnabled(True)

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
    progress_signal = pyqtSignal(int)
    status_signal = pyqtSignal(str)

    def __init__(self, cs, mode: list, cqc_info: list):
        super(checkinThread, self).__init__()
        self.cs = cs
        self.mode = mode
        self.cqc_info = cqc_info
        

    def run(self):
        if cs.checkActive():
            cqc_num, part_name, qty, cqe, pe, ins, event, cqc_type = self.cqc_info
            progress = 0
            taskqty = 1 + self.mode.count(False)
            if self.mode[3]:
                if event == 'RCT':
                    pass
                    #close RCT
                else:
                    pass
                    #close RCV
            
            if self.mode[2]:
                if self.cs.createAction(cqc_num, cqc_type, cqe):
                    if self.cs.createEvent(cqc_num, cqe, cqe, 'PRP', 'Sample cleaning'):
                        self.status_signal.emit('PRP created. ')
                        self.progress_signal.emit(int((progress+1)*100/taskqty))
                
                

            if self.mode[1]:
                if self.printLabel(cqc_num, part_name, cqe, pe, ins):
                    self.status_signal.emit('Label printed. ')
                    self.progress_signal.emit(int((progress+1)*100/taskqty))
            
            
            #record the handling
            pass
        else:
            self.progress_signal.emit(101)
            
    
    def printLabel(self, cqc_num, part_name, cqe, pe, ins):
        try: 
            style = PS('style', fontName="Helvetica-Bold", fontSize=8, leading=9, alignment=4)
            story = canvas.Canvas('label.pdf', (6*cm, 4*cm))
            story.setFont('Helvetica-Bold',9)
            story.drawString(0.2*cm, 3.6*cm, 'Product:')
            story.drawRightString(5.8*cm, 3.6*cm, part_name)
            story.drawString(0.2*cm, 3.6*cm-11, 'CQE:')
            story.drawRightString(5.8*cm, 3.6*cm-11, cqe)
            story.drawString(0.2*cm, 3.6*cm-22, 'PE:')
            story.drawRightString(5.8*cm, 3.6*cm-22, pe)
            p = Paragraph('Instruction: '+ins, style)
            x, y = p.wrap(5.6*cm, 18)
            p.drawOn(story, 0.2*cm, 3.6*cm-23-y)
            story.setFont('Helvetica-Bold',5)
            story.drawCentredString(3*cm, 3.6*cm-51, '---------------------------------------------COTS---------------------------------------------')
            Code39(cqc_num, ImageWriter(), False).save('bc', dict(text_distance=1.0, module_height=6, font_size=12, quiet_zone=1))
            story.drawImage('bc.png', 0.2*cm, 0*cm, width=5.6*cm, height=1.8*cm, preserveAspectRatio=True)
            os.remove('bc.png')
            story.save()
            args = 'gswin32c -dPrinted -dNOPAUSE -dBATCH -dNOSAFER -q -dNOPROMPT -dNumCopies=1 -dFitPage -sDEVICE=mswinpr2 ' \
                '-dDEVICEWIDTHPOINTS=170 ' \
                '-dDEVICEHEIGHTPOINTS=113 ' \
                '-sOutputFile="%printer%Deli DL-886A" ' \
                '"label.pdf"'
            subprocess.call(args, shell=True)
            os.remove('label.pdf')
            return True
        except:
            False

