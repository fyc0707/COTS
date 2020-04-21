#coding=utf-8
import datetime, re

import pandas as pd
import requests
from bs4 import BeautifulSoup

requests.urllib3.disable_warnings()

class CQCSniffer:
    
    headers = {
            'Accept' : 'image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, */*',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'en-GB',
            'User-Agent' : 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; Trident/4.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; wbx 1.0.0)',
            'Connection': 'Keep-Alive',
            'Cache-Control': 'no-cache'
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
        name = None
        try:
            resp = self.session.post(url+'login.do?method=login', data=loginParam, headers=self.headers, verify=False).text
            soup = BeautifulSoup(resp, 'html5lib')
            name = soup.find('b', text='Logged in Userid:')
        except:
            pass
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
                resp = self.session.get(self.url + turl, verify = False, headers=self.headers, timeout=700)
                s = False
                return resp
            except:
                self.session.get(self.url+'login.do?method=homepage', verify=False, headers=self.headers, timeout=700)
                return False


    def logOut(self):
        self.session.get(self.url+'login.do?method=logout')


    def getWIPData(self, fp):
        df = pd.read_excel(fp, header=7).astype(str)
        #df = df[df['Event Type'].str.contains('RCT|RCV')]
        df = df.drop(['2nd UD field Reception'], axis=1)
        df.columns = ['CQC#','Type','CQE','Customer','Part Name','Qty', 'Trace Code','Instruction', 'Event', 'B2B']
        df['B2B'] = df['B2B'].apply(lambda x: False if pd.isna(x) else True)
        df['Instruction'] = df['Instruction'].apply(lambda x: str(x)[19:])
        #df.reset_index(drop=True, inplace=True)
        return df

    def getBookmark(self, bookid, filename):
        self.session.get(self.url+'advancedSearch.do?method=advancedSearchBookMarkResults,bookid='+bookid, headers=self.headers, verify=False, timeout=700)
        resp = self.session.get(self.url+'advancedSearch.do?method=advancedSearchResultsExcelExport', headers=self.headers, verify=False, timeout=700)
        with open(filename, 'wb') as output:
            output.write(resp.content)
    
    
    def closeRCV(self, cqc_num, cqc_type, cqe, qty):
        pass

    def createAction(self, cqc_num, cqc_type, cqe:str):
        try:
            cqe = cqe.split('(')[-1].split(')')[0]
            self.tryUrl('login.do?method=loadProxy&rid='+cqe)
            data = {
                'arractid' : '',
                'arrPreDef' : '',
                'arrCustReq' : '',
                'aiNumber' : 1,
                'arrEvent' : '',
                'arrEventType' : '',
                'arrLineitem' : '',
                'arrItem' : 'newActionItem',
                'arrRcommitDate' : '',
                'arrTitle' : '',
                'arrDesc' : 'Sample cleaning',
                'arrType' : 'Problem Verification',
                'arrOwner' : cqe,
                'arrExpEff' : '',
                'arrInter1' : 'Y',
                'arrCommitDate' : datetime.date.today().strftime('%d-%b-%Y'),
                'arrSenddate' : '',
                'arrStatus' : 'NEW',
                'arrCustStatus' : '',
                'arrEndDate' : '',
                'arrValid' : '',
                'arrResult' : '',
                'arrMesEff' : ' ',
                'arrValDate' : ''
            }
            self.tryUrl('actionPlan.do?method=actionItemInformation&strIncidentNo='+cqc_num+'&strIncidentType='+cqc_type)
            self.session.post(self.url+'actionPlan.do?method=actionMultiActionitemInsert&IncNo='+cqc_num, data=data, headers=self.headers, verify=False)
            resp = self.tryUrl('actionPlan.do?method=actionItemInformation&strIncidentNo='+cqc_num).text
            soup = BeautifulSoup(resp, 'html5lib')
            if soup.find('textarea', text='Sample cleaning'):
                return True
            else:
                return False
        except:
            return False



    def createEvent(self, cqc_num, cqc_type, cqe, owner, event, comment):
        try:
            cqe = cqe.split('(')[-1].split(')')[0]
            owner = owner.split('(')[-1].split(')')[0]
            self.tryUrl('login.do?method=loadProxy&rid='+cqe)
            resp = self.tryUrl('getWorkFlowDetails.do?method=getEventsDetails&strIncidentNo='+cqc_num+'&strIncidentType='+cqc_type).text
            soup = BeautifulSoup(resp, 'html5lib')
            count = soup.find(id='lastIndexCount')['value']
            data = dict()
            for i in range(int(count)+1):
                field = [
                    'events['+str(i)+'].strEventNo',
                    'events['+str(i)+'].strEventType',
                    'events['+str(i)+'].strLineItem',
                    'events['+str(i)+'].strParentEvent',
                    'events['+str(i)+'].strParentName',
                    'events['+str(i)+'].strAssignToLocation',
                    'events['+str(i)+'].strRequestDate',
                    'events['+str(i)+'].strStatus',
                    'events['+str(i)+'].strAssignTo',
                    'events['+str(i)+'].strAssignToName',
                    'events['+str(i)+'].strDueDate',
                    'events['+str(i)+'].strInstructions',
                    'events['+str(i)+'].strStartDate',
                    'events['+str(i)+'].strComments',
                    'events['+str(i)+'].strCloseDate'
                ]
                for f in field:
                    value = soup.find(attrs={'name':f})
                    if value:
                        if value.has_attr('value'):
                            data[f] = value['value']
                        else:
                            data[f] = value.string
                    else:
                        data[f] = ''
            
            data['strPhase'] = soup.find('option', selected='selected')['value']
            data['strCQINo'] = cqc_num
            data['strIncidentType'] = cqc_type
            if soup.find(id='addEventButton'):
                resp = self.session.post(self.url+'getWorkFlowDetails.do?method=createEvent&strNewEventType='+event+
                '&strNewParentEvent='+data['strPhase']+
                '&strNewEventInstructions='+comment+
                '&strNewEventAssignTo='+owner+
                '&strNewEventDueDate=&strRCTNeeded=N&strNewLineItem=&strPAQConfirmation=', data=data, headers=self.headers, verify=False)
                if 'Event created successfully.' in resp.text:
                    return True
                else:
                    return False
            else:
                return False

        except:
            return False


    def closeEvent(self, cqc_num, cqc_type, cqe, event, comment):
        try:
            cqe = cqe.split('(')[-1].split(')')[0]
            self.tryUrl('login.do?method=loadProxy&rid='+cqe)
            resp = self.tryUrl('getWorkFlowDetails.do?method=getEventsDetails&strIncidentNo='+cqc_num+'&strIncidentType='+cqc_type).text
            soup = BeautifulSoup(resp, 'html5lib')
            count = soup.find(id='lastIndexCount')['value']
            data = dict()
            for i in range(int(count)+1):
                field = [
                    'events['+str(i)+'].strEventNo',
                    'events['+str(i)+'].strEventType',
                    'events['+str(i)+'].strLineItem',
                    'events['+str(i)+'].strParentEvent',
                    'events['+str(i)+'].strParentName',
                    'events['+str(i)+'].strAssignToLocation',
                    'events['+str(i)+'].strRequestDate',
                    'events['+str(i)+'].strStatus',
                    'events['+str(i)+'].strAssignTo',
                    'events['+str(i)+'].strAssignToName',
                    'events['+str(i)+'].strDueDate',
                    'events['+str(i)+'].strInstructions',
                    'events['+str(i)+'].strStartDate',
                    'events['+str(i)+'].strComments',
                    'events['+str(i)+'].strCloseDate'
                ]
                for f in field:
                    value = soup.find(attrs={'name':f})
                    if value:
                        if value.has_attr('value'):
                            data[f] = value['value']
                        else:
                            data[f] = value.string
                    else:
                        data[f] = ''
            
            data['strPhase'] = soup.find('option', selected='selected')['value']
            data['strCQINo'] = cqc_num
            data['strIncidentType'] = cqc_type
            index=''
            eventcp=''
            close_btn = soup.find_all(id='eventClose')
            for i in close_btn:
                index = i['onclick'].split("','")[-2]
                eventcp = i['onclick'].split("','")[-1].split("')")[0]
                if data['events['+index+'].strEventType'] == event:
                    break
                else:
                    index=''
                    eventcp=''
            
            if index:
                data['events['+str(index)+'].strComments'] = comment
                resp = self.session.post(self.url+'getWorkFlowDetails.do?method=closeEvent&index='+index+'&strEventCP='+eventcp, data=data, headers=self.headers, verify=False)
                if 'Event closed successfully.' in resp.text:
                    return True
                else:
                    return False
            else:
                return False
        
        except:
            return False
    

