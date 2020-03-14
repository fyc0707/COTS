import requests
import pandas as pd
import datetime
from bs4 import BeautifulSoup

requests.urllib3.disable_warnings()
pd.set_option('display.max_colwidth', None)

class CQCSniffer:
    
    headers = {
            'User-Agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36',
            'Accept' : 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
            'Accept-Language' : 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
            'Accept-Encoding': 'gzip, deflate, br'
        }

    def __init__(self, url, wbi, user_password):
        loginParam = {
            'strCoreId' : wbi,
            'strPage' : '',
            'strIncidentNo' : '',
            'strIncidentType' : '',
            'strCurruntPhase' : '',
            'strNotificationType' : '',
            'strAttchId' : '',
            'strCompId' : '',
            'strDuns' : '',
            'strPassword' : user_password
        }
        self.activeFlag = False
        self.url = url
        self.session = requests.Session()
        resp = self.session.post(url+'login.do?method=login', data=loginParam, headers=self.headers, verify=False).text
        soup = BeautifulSoup(resp, 'html5lib')
        name = soup.find('b', text='Logged in Userid:')
        if name:
            self.user_name = name.parent.parent.next_sibling.next_sibling.attrs['title']
            self.activeFlag = True
        else:
            pass

    def checkActive(self):
        try:
            resp = self.session.get(self.url+'login.do?method=homepage', verify=False, headers=self.headers, timeout=700).text
            soup = BeautifulSoup(resp, 'html5lib')
            name = soup.find('b', text='Logged in Userid:')
            if name:
                self.activeFlag = True
                return True
            else:
                self.activeFlag = False
                return False
        except:
            self.activeFlag = False
            return False

        soup = BeautifulSoup(resp, 'html5lib')
        name = soup.find('b', text='Logged in Userid:')
        if name:
            self.activeFlag = True
            return True
        else:
            self.activeFlag = False
            return False


    def tryUrl(self, turl):
        f = True
        while f:
            try:
                self.session.get(self.url + turl, verify = False, headers=self.headers, timeout=700)
                s = False
            except:
                self.session.get(self.url+'login.do?method=homepage', verify=False, headers=self.headers, timeout=700)
        
    def logOut(self):
        self.session.get(self.url+'login.do?method=logout')
    
    def getRCVData(self, bookid):
        self.session.get(self.url+'advancedSearch.do?method=advancedSearchBookMarkResults&bookid='+bookid, headers=self.headers, verify=False, timeout=700)
        resp = self.session.get(self.url+'advancedSearch.do?method=advancedSearchResultsExcelExport', headers=self.headers, verify=False, timeout=700)
        with open('output.xlsx', 'wb') as output:
            output.write(resp.content)
        wb = xl.load_workbook('output.xlsx')
        ws = wb.active
        ws.delete_rows(1,7)
        data = ws.values
        cols = next(data)[0:]
        data = list(data)
        data = (islice(r,0,None) for r in data)
        self.dataframe = pd.DataFrame(data, columns=cols)
        print(self.dataframe)

    def getBookmark(self, bookid, filename):
        self.session.get(self.url+'advancedSearch.do?method=advancedSearchBookMarkResults&bookid='+bookid, headers=self.headers, verify=False, timeout=700)
        resp = self.session.get(self.url+'advancedSearch.do?method=advancedSearchResultsExcelExport', headers=self.headers, verify=False, timeout=700)
        with open(filename, 'wb') as output:
            output.write(resp.content)
    
    def getReportDrumbeat(self, window_size=2):
        df = self.dataframe[self.dataframe['Complaint Status'].isin(['IP','OP'])]
        df = df[df['Business Line'].isin(['BLC1','BLC2','BLC3','BLC4','BLSP','BLAU'])]
        df.reset_index(inplace=True)
        RDdataframe = pd.DataFrame(columns=['CQC#','CQE','Customer','Flagged','Due Report','Due Time'])
        due_time = []
        for index, rows in df.iterrows():
            init_time = rows['Complaint Start Date']
            rcv_time = rows['Receive Date']
            last_comm_time = None
            last_comm = None
            if rows['Customer Complaint ID'] != '':
                last_comm_time = self.getB2BLastComm(rows['Complaint ID'])
            else:
                last_comm_time = rows['Initial Send Date']
                try:
                    last_comm_time = datetime.datetime.strftime(last_comm_time, '%m-%d-%Y %H:%M')
                except:
                    pass
                last_comm = rows['Communication Type']
                if last_comm_time != '':
                    receive = last_comm_time.split(',').count('Receive')
                    comm_number = len(last_comm_time.split(','))
                    last_comm_time = last_comm_time.split(',')[-1]
                    last_comm_time = datetime.datetime.strptime(last_comm_time, '%m-%d-%Y %H:%M')
                    last_comm = last_comm.split(',')[comm_number-1]
                    if comm_number-receive == 0:
                        last_comm = None
                        last_comm_time = None
                else:
                    last_comm_time = None
            RDdataframe.loc[index,'CQC#'] = rows['Complaint ID']
            RDdataframe.loc[index,'CQE'] = rows['CQE']
            RDdataframe.loc[index,'Customer'] = rows['Logical Customer Name']
            RDdataframe.loc[index,'Flagged'] = rows['Customer Responsive Flag']
        
            
            #B2B Communications
            
            if last_comm_time == None:
                RDdataframe['Due Report'].loc[index] = 'FPC'
                RDdataframe['Due Time'].loc[index] = init_time + datetime.timedelta(2,0,0) + self.tz_offset
                due_time.append(init_time + datetime.timedelta(2,0,0))
            elif rcv_time != None:
                if last_comm_time == None:
                        RDdataframe['Due Report'].loc[index] = 'Initial'
                        RDdataframe['Due Time'].loc[index] = rcv_time + datetime.timedelta(2,0,0) + self.tz_offset
                        due_time.append(rcv_time + datetime.timedelta(2,0,0))
                else:
                    if (last_comm_time + datetime.timedelta(5,0,0) > rcv_time) and (last_comm_time < rcv_time):
                        RDdataframe['Due Report'].loc[index] = 'Initial'
                        RDdataframe['Due Time'].loc[index] = rcv_time + datetime.timedelta(2,0,0) + self.tz_offset
                        due_time.append(rcv_time + datetime.timedelta(2,0,0))
                    else:
                        RDdataframe['Due Report'].loc[index] = 'Interim'
                        RDdataframe['Due Time'].loc[index] = last_comm_time + datetime.timedelta(7,0,0) + self.tz_offset
                        due_time.append(last_comm_time + datetime.timedelta(7,0,0))
            else:
                RDdataframe['Due Report'].loc[index] = 'Interim'
                RDdataframe['Due Time'].loc[index] = last_comm_time + datetime.timedelta(7,0,0) + self.tz_offset
                due_time.append(last_comm_time + datetime.timedelta(7,0,0))
            
        cqc_to_drop = []
        current_time = datetime.datetime.now() - self.tz_offset
        coverage = 2
        if datetime.datetime.today().weekday() == 4:
            coverage = 3

        #Clean the dataframe
        for i in range(len(RDdataframe)):
            if due_time[i]==None:
                cqc_to_drop.append(i)
            elif due_time[i]-current_time > datetime.timedelta(coverage+0.01,0,0):
                cqc_to_drop.append(i)
        RDdataframe.drop(cqc_to_drop, inplace=True)
        RDdataframe.sort_values(by='Due Time', inplace=True)
        RDdataframe.reset_index(drop=True, inplace=True)
        RDdataframe['Due Time'] = RDdataframe['Due Time'].apply(datetime.datetime.strftime, args=('%b-%d %I:%M %p',))
        print(RDdataframe)
        return(RDdataframe)

    def getBacklogReport(self):
        df = self.dataframe[self.dataframe['Business Line'].isin(['BLC1','BLC2','BLC3','BLC4','BLSP','BLAU'])]
        df.reset_index(inplace=True)
        drop_col = []
        columns = df.columns.tolist()
        for i in [0,7,8,9,10,11,13,16,18,19,22,23,24,25,28,29,30,31,32,36,37,38]:
            drop_col.append(columns[i])
        df.drop(columns=drop_col, inplace=True)
        print(df)
        return(df)
    
    def getCARReport(self):
        df = self.dataframe[self.dataframe['Business Line'].isin(['BLC1','BLC2','BLC3','BLC4','BLSP','BLAU'])]
        CARdataframe = pd.DataFrame(columns=['CQC#', 'CQE', 'Part Type', 'Date Code', 'Customer', 'CAR Owner', 'Event CT', 'Occur Code', 'Occur DESC.', 'FA Code', 'FA DESC.', 'Comments'])
        i = 0
        for index, row in df.iterrows():
            if row['Corrective Action #'] == '':
                continue
            else:
                t = row['Event Type'].split(',').index('CAR')
                t = str(row['Event IP Cycle Time']).split(',')[t]
                CARdataframe.loc[i] = [row['Complaint ID'], row['CQE'], row['Part Type Name'], row['Assembly Marked Date Code'], row['Logical Customer Name'], row['CAR Event Owner'], float(t), row['Fail Mech Occur Code'], row['Fail Mech Occur Desc'], row['Failure Code'], row['FAILURE CODE DESCRIPTION'], ' ']
                i = i + 1
        CARdataframe.sort_values('Event CT', ascending=False, inplace=True)
        CARdataframe.reset_index(drop=True, inplace=True)
        print(CARdataframe)
        return CARdataframe
    
    def getRWReport(self):
        df = self.dataframe[self.dataframe['Business Line'].isin(['BLC1','BLC2','BLC3','BLC4','BLSP','BLAU'])]
        RWdataframe = pd.DataFrame(columns=['CQC#', 'CQE', 'Part Type', 'Customer', 'Event', 'Event CT', 'Comments'])
        i = 0
        for index, row in df.iterrows():
            if 'PRW' in row['Event Type']:
                t = row['Event Type'].split(',').index('PRW')
                t = str(row['Event IP Cycle Time']).split(',')[t]
                RWdataframe.loc[i] = [row['Complaint ID'], row['CQE'], row['Part Type Name'], row['Logical Customer Name'], 'PRW', float(t), ' ']
                i = i + 1
            elif 'CAW' in row['Event Type']:
                t = row['Event Type'].split(',').index('CAW')
                t = str(row['Event IP Cycle Time']).split(',')[t]
                RWdataframe.loc[i] = [row['Complaint ID'], row['CQE'], row['Part Type Name'], row['Logical Customer Name'], 'CAW', float(t), ' ']
                i = i + 1
            else:
                continue
        RWdataframe.sort_values('Event CT', ascending=False, inplace=True)
        RWdataframe.reset_index(drop=True, inplace=True)
        print(RWdataframe)
        return RWdataframe

    def getICETSTReport(self):
        df = self.dataframe[self.dataframe['Business Line'].isin(['BLC1','BLC2','BLC3','BLC4','BLSP','BLAU'])]
        ICEdataframe = pd.DataFrame(columns=['CQC#', 'CQE', 'Part Type', 'Customer', 'Event', 'Event CT', 'Comments', 'BL'])
        i = 0
        for index, row in df.iterrows():
            if 'TST' in row['Event Type']:
                t = row['Event Type'].split(',').index('TST')
                t = str(row['Event IP Cycle Time']).split(',')[t]
                ICEdataframe.loc[i] = [row['Complaint ID'], row['CQE'], row['Part Type Name'], row['Logical Customer Name'], 'TST', float(t), ' ', row['Business Line']]
                i = i + 1
            elif 'ICE' in row['Event Type']:
                t = row['Event Type'].split(',').index('ICE')
                t = str(row['Event IP Cycle Time']).split(',')[t]
                ICEdataframe.loc[i] = [row['Complaint ID'], row['CQE'], row['Part Type Name'], row['Logical Customer Name'], 'ICE', float(t), ' ', row['Business Line']]
                i = i + 1
            else:
                continue
        
        ICEdataframe.sort_values('Event CT', ascending=False, inplace=True)
        ICEdataframe.reset_index(drop=True, inplace=True)
        print(ICEdataframe)
        return ICEdataframe

    def getB2BLastComm(self, cqc_number):
        last_comm_time = None
        for row in self.tdsession.execute('select SUBMIT_DATETIME from "EDW"."CQC_B2B_COMMUNICATION" where INCIDENT_NUM=\''+cqc_number+'\' and PROBLEM_SOLVER_SUMMARY_INDICATOR=\'S\';'):
            if last_comm_time == None:
                last_comm_time = row.values[0]
            if row.values[0]>last_comm_time:
                last_comm_time = row.values[0]
        return last_comm_time
