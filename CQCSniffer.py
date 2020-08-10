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
        self.url = url
        self.wbi = wbi
        self.user_name = ''
        self.user_password = user_password
        self.activeFlag = False
        self.session = requests.Session()

    def login(self):
        loginParam = {
            'strCoreId' : self.wbi,
            'strPage' : '',
            'strIncidentNo' : '',
            'strIncidentType' : '',
            'strCurruntPhase' : '',
            'strNotificationType' : '',
            'strAttchId' : '',
            'strCompId' : '',
            'strDuns' : '',
            'strPassword' : self.user_password
        }
        
        name = None
        try:
            resp = self.session.post(self.url+'login.do?method=login', data=loginParam, headers=self.headers, verify=False).text
            soup = BeautifulSoup(resp, 'html5lib')
            name = soup.find('b', text='Logged in Userid:')
        except Exception as err:
            print(err)
            return False
        if name:
            self.user_name = name.parent.parent.next_sibling.next_sibling.attrs['title']
            self.activeFlag = True
            return True
        else:
            return False

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
        df = pd.read_excel(fp, header=7, keep_default_na=False).astype(str)
        for i, row in df.iterrows():
            if row['Part Type Name']=='':
                row['Part Type Name'] = row['Part Type']
        df = df.drop(['Part Type'], axis=1)
        df = df.drop(['2nd UD field Reception'], axis=1)
        df.columns = ['CQC#','Type','CQE','Customer','Part Name','Qty', 'Trace Code', 'Ship Ref.','Instruction', 'Event', 'B2B']
        df['B2B'] = df['B2B'].apply(lambda x: False if pd.isna(x) else True)
        df['Instruction'] = df['Instruction'].apply(lambda x: str(x)[19:])
        return df

    def getBookmark(self, bookid, filename):
        self.session.get(self.url+'advancedSearch.do?method=advancedSearchBookMarkResults,bookid='+bookid, headers=self.headers, verify=False, timeout=700)
        resp = self.session.get(self.url+'advancedSearch.do?method=advancedSearchResultsExcelExport', headers=self.headers, verify=False, timeout=700)
        with open(filename, 'wb') as output:
            output.write(resp.content)
    
    
    def closeRCV(self, cqc_num, cqc_type, b2b, cqe, qty):
        try:
            self.tryUrl('login.do?method=loadProxy&rid='+'NXA08198')
            resp = self.tryUrl('getSmrylnItm.do?method=getSummryLineItems&strIncidentNo='+cqc_num+'&strIncidentType='+cqc_type).text
            soup = BeautifulSoup(resp, 'html5lib')
            qty = qty.split(',')
            if soup.find(id='lineitemrcv'):
                count = soup.find(id='lastIndexCount')['value']
                data = dict()
                for i in range(int(count)):
                    i = str(i)
                    field = [
                        'lineitemforms['+i+'].strfailNo',
                        'lineitemforms['+i+'].strIncidentType',
                        'lineitemforms['+i+'].strEditAccess',
                        'lineitemforms['+i+'].strPhase',
                        'lineitemforms['+i+'].strLineItemId',
                        'lineitemforms['+i+'].strTNIAccess',
                        'lineitemforms['+i+'].strLineItemDate',
                        'lineitemforms['+i+'].strLineItemNo',
                        'lineitemforms['+i+'].strIncidentNo',
                        'lineitemforms['+i+'].strdelchk',
                        'lineitemforms['+i+'].strrcvchk',
                        'lineitemforms['+i+'].strLineItemType',
                        'lineitemforms['+i+'].strCustReason',
                        'lineitemforms['+i+'].strABADone',
                        'lineitemforms['+i+'].strABAComments',
                        'lineitemforms['+i+'].strFqeCustDesc',
                        'lineitemforms['+i+'].strFqeComments',
                        'lineitemforms['+i+'].strCqeCustDesc',
                        'lineitemforms['+i+'].strCqeComments',
                        'lineitemforms['+i+'].strCustRef',
                        'lineitemforms['+i+'].strPrototype',
                        'lineitemforms['+i+'].strMileage',
                        'lineitemforms['+i+'].strEndCustCode',
                        'lineitemforms['+i+'].strCustBuildLoc',
                        'lineitemforms['+i+'].strECUManufacturingDate',
                        'lineitemforms['+i+'].strTraceCode',
                        'lineitemforms['+i+'].strSysCdeInit',
                        'lineitemforms['+i+'].strSysCdeValDB',
                        'lineitemforms['+i+'].strSysCdeDrpInit',
                        'lineitemforms['+i+'].strSysCdeDrpDb',
                        'lineitemforms['+i+'].strSysCdeTxtInit',
                        'lineitemforms['+i+'].strSysCdeTxtDb',
                        'lineitemforms['+i+'].strReturnCode',
                        'lineitemforms['+i+'].strFirstNo',
                        'lineitemforms['+i+'].strFirstYes',
                        'lineitemforms['+i+'].strSecondNo',
                        'lineitemforms['+i+'].strSecondYes',
                        'lineitemforms['+i+'].strUserAct',
                        'lineitemforms['+i+'].strRequestId',
                        'lineitemforms['+i+'].strRequestStatus',
                        'lineitemforms['+i+'].strSysCodeTxt',
                        'lineitemforms['+i+'].strEmptyCol',
                        'lineitemforms['+i+'].strSysCodeDisp',
                        'lineitemforms['+i+'].strDateCode',
                        'lineitemforms['+i+'].strTestMarkedDateCode',
                        'lineitemforms['+i+'].strBackMark',
                        'lineitemforms['+i+'].strMaskSet',
                        'lineitemforms['+i+'].strLineItemQty',
                        'lineitemforms['+i+'].strCustSerialNo',
                        'lineitemforms['+i+'].strLineItemComm',
                        'lineitemforms['+i+'].strFabSite',
                        'lineitemforms['+i+'].strWaferLotNo',
                        'lineitemforms['+i+'].strFabOutDate',
                        'lineitemforms['+i+'].strProbeSite',
                        'lineitemforms['+i+'].strProbeDate',
                        'lineitemforms['+i+'].strAssySite',
                        'lineitemforms['+i+'].strAssyLotNo',
                        'lineitemforms['+i+'].strAssyOutDate',
                        'lineitemforms['+i+'].strTestSite',
                        'lineitemforms['+i+'].strTestLotNo',
                        'lineitemforms['+i+'].strTestOutDate',
                        'lineitemforms['+i+'].strQtyRcvd',
                        'lineitemforms['+i+'].strFuncSafetyIssue',
                        'lineitemforms['+i+'].strCustFunSafetyRelIssue'
                    ]
                    if b2b:
                        field = field + [
                            'lineitemforms['+i+'].strLineCompId',
                            'lineitemforms['+i+'].strLineCompDuns',
                            'lineitemforms['+i+'].strAssignBtnAccess',
                            'lineitemforms['+i+'].strResetBtnAccess',
                            'lineitemforms['+i+'].strLineCompStatus',
                            'lineitemforms['+i+'].strLineCompDesc',
                            'lineitemforms['+i+'].strRcvInfo'
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
                    data['lineitemforms['+i+'].strQtyRcvd'] = str(int(qty[int(i)]))
                
                data['CANumber'] = soup.find(attrs={'name':'CANumber'})['value']
                data['strUserId'] = soup.find(attrs={'name':'strUserId'})['value']
                data['strPassword'] = soup.find(attrs={'name':'strPassword'})['value']
                data['strProxyId'] = soup.find(attrs={'name':'strProxyId'})['value']
                data['strPartNum'] = ''
                data['strSysCodeAvab'] = ''
                data['strSysCodeExist'] = ''
                data['strIncidentNo'] = soup.find(attrs={'name':'strIncidentNo'})['value']
                data['strIncidentType'] = soup.find(attrs={'name':'strIncidentType'})['value']
                data['strStatusCode'] = soup.find(attrs={'name':'strStatusCode'})['value']
                data['strCheckAdmin'] = soup.find(attrs={'name':'strCheckAdmin'})['value']
                data['strCheckFqe'] = soup.find(attrs={'name':'strCheckFqe'})['value']
                data['strCheckCqe'] = soup.find(attrs={'name':'strCheckCqe'})['value']
                data['strCheckReceptionCenter'] = soup.find(attrs={'name':'strCheckReceptionCenter'})['value']
                self.session.post(self.url+'getSmrylnItm.do?method=receiveLineitem', data=data, headers=self.headers, verify=False)
                resp = self.tryUrl('getSmrylnItm.do?method=getSummryLineItems&strIncidentType='+cqc_type+'&strIncidentNo='+cqc_num)
                if '<input type="hidden" name="lineitemforms[0].strQtyRcvd"' in resp.text:
                    return True
                else:
                    return False
            else:
                return False

        except Exception as err:
            print(err)
            return False

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
        except Exception as err:
            print(err)
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
                i = str(i)
                field = [
                    'events['+i+'].strEventNo',
                    'events['+i+'].strEventType',
                    'events['+i+'].strLineItem',
                    'events['+i+'].strParentEvent',
                    'events['+i+'].strParentName',
                    'events['+i+'].strAssignToLocation',
                    'events['+i+'].strRequestDate',
                    'events['+i+'].strStatus',
                    'events['+i+'].strAssignTo',
                    'events['+i+'].strAssignToName',
                    'events['+i+'].strDueDate',
                    'events['+i+'].strInstructions',
                    'events['+i+'].strStartDate',
                    'events['+i+'].strComments',
                    'events['+i+'].strCloseDate'
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
            
            data['strPhase'] = 'EVAL'#soup.find('option', selected='selected')['value']
            data['strCQINo'] = cqc_num
            data['strIncidentType'] = cqc_type
            if soup.find(id='addEventButton'):
                resp = self.session.post(self.url+'getWorkFlowDetails.do?method=createEvent&strNewEventType='+event+
                '&strNewParentEvent='+data['strPhase']+
                '&strNewEventInstructions='+comment+
                '&strNewEventAssignTo='+owner+
                '&strNewEventDueDate=&strRCTNeeded=N&strNewLineItem=&strPAQConfirmation=', data=data, headers=self.headers, verify=False)
                if 'Event created successfully.' in resp.text or comment in resp.text:
                    return True
                else:
                    return False
            else:
                return False

        except Exception as err:
            print(err)
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
                i = str(i)
                field = [
                    'events['+i+'].strEventNo',
                    'events['+i+'].strEventType',
                    'events['+i+'].strLineItem',
                    'events['+i+'].strParentEvent',
                    'events['+i+'].strParentName',
                    'events['+i+'].strAssignToLocation',
                    'events['+i+'].strRequestDate',
                    'events['+i+'].strStatus',
                    'events['+i+'].strAssignTo',
                    'events['+i+'].strAssignToName',
                    'events['+i+'].strDueDate',
                    'events['+i+'].strInstructions',
                    'events['+i+'].strStartDate',
                    'events['+i+'].strComments',
                    'events['+i+'].strCloseDate'
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
                if 'Event closed successfully.' in resp.text or comment in resp.text :
                    return True
                else:
                    return False
            else:
                return False
        
        except Exception as err:
            print(err)
            return False
    
    def getProductName(self, cqc_num):
        try:
            resp = self.session.post(self.url+'advancedSearch.do?method=advancedSearchIncidents', data = {'incidentNo' : cqc_num}, headers=self.headers, verify=False).text
            soup = BeautifulSoup(resp, 'html5lib')
            name = soup.find(id='strLogicPartName')['value']
            return str(name)
            
        except Exception as err:
            print(err)
            return None

    def getCQEName(self, cqc_num):
        try:
            resp = self.session.post(self.url+'advancedSearch.do?method=advancedSearchIncidents', data = {'incidentNo' : cqc_num}, headers=self.headers, verify=False).text
            soup = BeautifulSoup(resp, 'html5lib')
            name = str(soup.find(attrs={'name':'strCQENameDesc'})['value'])
            wbi = name.split('(')[-1].split(')')[0]
            name = name.split('(')[-2]
            name = name+' ('+wbi+')'
            return str(name)
            
        except Exception as err:
            print(err)
            return None

    def getEmail(self, name):
        try:
            name = name.split('(')[-1].split(')')[0]
            resp = self.session.get(self.url+'login.do?method=getOtherProfile&rid='+name, verify=False, headers=self.headers)
            soup = BeautifulSoup(resp.text, 'html5lib')
            email = soup.find('b', text='Department').parent.parent.previous_sibling.previous_sibling.a.next_element
            return email

        except Exception as err:
            print(err)
            return None
        