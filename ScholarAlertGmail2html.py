from __future__ import print_function
import numpy as np
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from apiclient import errors
import yaml
from bs4 import BeautifulSoup
from bs4 import Tag
import base64
from retrying import retry
import math
import re
from datetime import datetime, date, timedelta
import csv
from pytz import timezone
import copy
from pytictoc import TicToc
import shelve
import bibtexparser
from bibtexparser.bparser import BibTexParser
from bibtexparser.customization import convert_to_unicode
from bibtexparser.bibdatabase import BibDatabase
from bibtexparser.bwriter import BibTexWriter
import os
from os import path as ospath
from os import makedirs
import webbrowser


# Use http proxy as a global proxy
# os.environ["http_proxy"] = "http://127.0.0.1:10809"
# os.environ["https_proxy"] = "http://127.0.0.1:10809"
file_dir = os.path.dirname(__file__)
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
user_id = 'me'
scholar_email = 'scholaralerts-noreply@google.com'
cst_tz = timezone('Asia/Shanghai')
BROWSER_COMMAND = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
subType_AuthCitation = 'citation'
subType_AuthNew = 'new'
subType_AuthRelated = 'related'

t = TicToc()

@retry(wait_random_min=100, wait_random_max=1000, stop_max_attempt_number=4)
# Takes a message id and reads the message using google api
def readMessage(gmail, message_id, format="full"):
    return gmail.users().messages().get(id=message_id, userId=user_id, format=format).execute()

# pulls date string from the header
def getDate(headers):
    for header in headers:
        if header['name'] == 'Date' or header['name'] == 'date':
            dtstr = header['value']
            try:
                datetimeMsg = datetime.strptime(
                    re.findall("^\w+, \d+ \w+ \d+ \w+:\w+:\w+ -?\d+",dtstr)[0],
                    '%a, %d %b %Y %H:%M:%S %z').astimezone(cst_tz)
            except:
                datetimeMsg = datetime.strptime(
                    re.findall("^\w+, \d+ \w+ \d+ \w+:\w+:\w+",dtstr)[0],
                    '%a, %d %b %Y %H:%M:%S').astimezone(cst_tz)
    return datetimeMsg

# pulls date string from the header
def getEmailFrom(scholarMessage):
    headers = scholarMessage['payload']['headers']
    for header in headers:
        if header['name'] == 'From' or header['name'] == 'from':
            return header['value']

# The send date in scholarMessage from now
def daysMsgFromNow(scholarMessage):
    dateNow = datetime.now().astimezone(cst_tz)
    datetimeMsg = getDate(scholarMessage['payload']['headers'])
    return (dateNow - datetimeMsg).days

# get all the messages using messagesID
# save them as pkl, as well as save the emails as html files
@retry(wait_random_min=100, wait_random_max=1000, stop_max_attempt_number=4)
def pullMessage(gmail, messages, maxDays, maxRange, scholarMessages):
    # scholarMessages = list()
    sglMsgDict = mkMsgDict(scholarMessages)
    flagStop = True
    idxRange = [0, maxRange]
    idx_start = idxRange[0]
    idx_now = idx_start
    # for idx in range(idxRange[0], idxRange[1]):  # a in messages:
    while flagStop:
        message_id = messages[idx_now]['id']
        if not message_id in sglMsgDict:
            print('new msg %d' % idx_now)
            sglMsg = readMessage(gmail, message_id)
            scholarMessages.append(sglMsg)
            print('%d-days %d' % (idx_now+1, daysMsgFromNow(sglMsg)))
            if daysMsgFromNow(sglMsg)> maxDays:
                flagStop = False
        if idx_now == maxRange - 1:
            flagStop = False
        if math.fmod(idx_now+1, 200) == 0 or \
                idx_now == maxRange - 1 or \
                flagStop == False:
            print('%d' %(idx_now+1))
        idx_now += 1

        # # save scholarMessages as pkl when idx is the multiple of 50
        # if math.fmod(idx_now+1, 100) == 0 or \
        #         idx_now == maxRange - 1 or \
        #         flagStop == False:
        #     fileNamePkl = 'pkl/scholarMessages/M_' + \
        #                   str(idx_start+1) + '_' + messages[idx_start]['id'] + '_to_' + \
        #                   str(idx_now+1) + '_' + message_id + '.pkl'
        #     with open(fileNamePkl, 'wb') as fmsg:
        #         pickle.dump(scholarMessages, fmsg)
    scholarMessages = sorted(scholarMessages, key=lambda k: k['id'], reverse=True)
    return scholarMessages


# Parsing the scholarMessages
# Converting msg to pub
# use forceRead = True when del the publication and re_read all scholarMessages, during this,
# the scholarMessages[idx]['readFlag'] is clear
def msg2Pub(scholarMessages, publications, forceRead = False):
    if len(publications) == 0:
        forceRead = True
    if forceRead:
        publications = list()
        for i in range(len(scholarMessages)):
            if 'readFlag' in scholarMessages[i]:
                scholarMessages[i]['readFlag'] = False
    pubTitDict = mkPubTitDict(publications)
    for idx in range(len(scholarMessages)):
        sMsg = scholarMessages[idx]
        if 'readFlag' not in sMsg or forceRead or sMsg['readFlag'] == False:
            if getEmailFrom(sMsg).find(scholar_email) < 0:
                continue
            body = base64.urlsafe_b64decode(sMsg['payload']['body']['data'])
            soup: BeautifulSoup = BeautifulSoup(body, 'html.parser')
            # nPub = len(soup.body.find_all('h3'))
            tagPubs = soup.body.find_all('h3')
            for tagPub in tagPubs:
                title = getTitle(tagPub)
                if not title in pubTitDict:
                    pubTitDict[title] = len(publications)
                    publications.append(Publication(tagPub))
                publications[pubTitDict[title]].addHeaders(sMsg)
            if math.fmod(idx+1, 200) == 0:
                print('msg2Pub of idx = %d' %(idx+1))
        scholarMessages[idx]['readFlag'] = True
    return [publications, pubTitDict, idx + 1]

# get subject from the header
def getSubject(headers):
    for header in headers:
        if header['name'] == 'Subject':
            return header['value']

# Change special characters in a string into spaces to become a legal file name
def correct_FileName(strName, str1=' '):
    error_set = ['/', '\\', ':', '*', '?', '"', '|', '<', '>']
    for c in strName:
        if c in error_set:
            strName = strName.replace(c, str1)
    return strName

# get title string from a raw paper
def getTitle(pubTag):
    # title is in the first tag under 'a'
    if pubTag.find('a'):
        return pubTag.a.text
    else:
        return pubTag.text

    return

# Get the service interface of GmailApi, validate and save
# the token.json and credentials.json files of GmailAPI
@retry(wait_random_min=100, wait_random_max=1000, stop_max_attempt_number=4)
def getGmailApi():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if not ospath.exists('json/'):
        makedirs('json/')
    if os.path.exists('json/token.json'):
        creds = Credentials.from_authorized_user_file('json/token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'json/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('json/token.json', 'w') as token:
            token.write(creds.to_json())
    try :
        gmail = build('gmail', 'v1', credentials=creds)
        return gmail
    except errors.HttpError as error:
        print('An error occurred: %s' % error)
        raise error

# Define the Publication class to save publications,
# including bib publication information,
# the publication content of the original email in the soup type,
# the list of subjects corresponding to the publication,
# the date and time list of the email,
# and the rating of the publication
class Publication(object):
    def __init__(self, pubTag):
        titleTag = pubTag
        authorTag = pubTag.next_sibling
        abstractTag = authorTag.next_sibling
        shareTag = abstractTag.find_next_sibling('div')

        self.bib = dict()
        title = getTitle(titleTag)
        self.bib['title'] = title

        urlstr = titleTag.find('a')['href']
        self.bib['url'] = urlstr
        # urlstrip = re.findall('(?<=url=).*?(?=&)', urlstr)
        # if len(urlstrip) > 0:
        #     self.bib['url'] = urlstrip[0]
        # else:
        #     self.bib['url'] = urlstr

        authorYearStr = authorTag.text.replace(u'\xa0', u' ')
        self.bib['author'] = ' and '.join([i.strip() for i in authorYearStr.split(' - ')[0].split(',')])
        if len(authorYearStr.split(' - '))>1:
            jonlYearStr = authorYearStr.split(' - ')[1]
            year = re.findall('\\d{4}', jonlYearStr)
            if len(year) > 0:
                self.bib['year'] = year[0]
                jonlYearStr = jonlYearStr.replace(year[0], '')
            journal = jonlYearStr.split(',')
            if len(journal) > 0:
                self.bib['journal'] = journal[0]

        if 'class' in abstractTag.attrs\
                and abstractTag.attrs['class'][0] == 'gse_alrt_sni':
            self.bib['abstract'] = abstractTag.text
        self.subjects = []
        self.messageIDs = []
        self.dateLists = []
        self.score = 0

        htmlstr = ' <html xmlns="http://www.w3.org/1999/xhtml" ' \
                  'xmlns:o="urn:schemas-microsoft-com:office:office">' \
                  '<head><style>body{background-color:#fff}.gse_alrt_title{text-decoration:none}' \
                  '.gse_alrt_title:hover{text-decoration:underline} @media screen and (max-width: 599px) ' \
                  '{.gse_alrt_sni br{display:none;}}</style></head>' \
                  '<body><div style="font-family:arial,sans-serif;' \
                  'font-size:13px;line-height:16px;color:#222;width:100%;max-width:600px"/></div></body></html>'
        self.soup = BeautifulSoup(htmlstr, 'html.parser')
        # tag = self.soup.html.body.div
        # new_div = self.soup.new_tag('pub')
        # tag.append(new_div)
        # tag = self.soup.html.body.div.pub
        # if len(titleTag.a['href'])>0:
        #     titleTag.a['href'].replace(self.bib['url'])
        self.soup.html.body.div.append(titleTag)
        self.soup.html.body.div.append(authorTag)
        self.soup.html.body.div.append(abstractTag)
        self.soup.html.body.insert(1, shareTag)
        # https://blog.csdn.net/lu8000/article/details/82313054
        # saveSoupTag(self.soup)
        # a = self.soup.prettify()

    def addHeaders(self, sMsg):
        headers = sMsg['payload']['headers']
        subject = getSubject(headers)
        datetimMsg = getDate(headers)
        self.subjects.append(subject)
        self.messageIDs.append(sMsg['id'])
        self.dateLists.append(datetimMsg)

    def ratingScore(self, authVal, jonlVal):
        typeVal = dict()
        typeVal[subType_AuthCitation] = 2
        typeVal[subType_AuthNew] = 8
        typeVal[subType_AuthRelated] = 1
        typeVal[''] = 0.5
        authType = pubSub2AuthorType(self.subjects)
        rateScore = 0
        for aT in authType:
            authScore = 1
            typeScore = 1
            if aT[0] in authVal:
                authScore = authVal[aT[0]]
            if aT[0] in authVal:
                typeScore = authVal[aT[0]]
            rateScore += authScore*typeScore
        jonlScore = 1.0
        if 'journal' in self.bib:
            if self.bib['journal'] in jonlVal:
                jonlScore = jonlVal[self.bib['journal']]
        rateScore = rateScore*jonlScore
        self.score = rateScore
        return self.score

# rate the publications and get the soted scores and the sorted idxe
def rateSortPubs(publications, authVal, jonlVal):
    scorePubs = [0.0 for i in range(len(publications))]
    for idx in range(len(publications)):
        scorePubs[idx] = publications[idx].ratingScore(authVal, jonlVal)
    sorted_idx = sorted(range(len(scorePubs)),
                       key=lambda k: publications[k].score,
                       reverse=True)
    sorted_scorePubs = sorted(scorePubs, reverse=True)
    return [scorePubs, sorted_scorePubs, sorted_idx]

# save the pub to html file according to date range
def savPub2html(publications, sorted_idx, fileNameHtml = 0, dateRange=0):
    if dateRange == 0:
        dateRange = [date.today() - timedelta(days=30), date.today()]
    trueDateTimeRange = [datetime.now()]*2
    pubDatetimeMin = list()
    pubDatetimeMax = list()
    pub2html = list()

    idx_pub = 0
    for i in range(len(sorted_idx)):
        pub = publications[sorted_idx[i]]
        if (min(pub.dateLists).date() - dateRange[0]).days >= 0 and \
                (min(pub.dateLists).date() - dateRange[1]).days <= 0:
            pub2html.append(pub)
            pubDatetimeMin.append(min(pub.dateLists))
            pubDatetimeMax.append(max(pub.dateLists))
    trueDateTimeRange[0] = min(pubDatetimeMin)
    trueDateTimeRange[1] = max(pubDatetimeMax)
    # html head
    htmlstr = ' <html xmlns="http://www.w3.org/1999/xhtml" xmlns:o="urn:schemas-microsoft-com:office:office">' \
              '<head> <meta http-equiv="Content-Type" content="text/html;charset=UTF-8" /> <style>' \
              '.main{ text-align: left;  background-color: #fff; margin: auto; ' \
              'position: absolute; top: 110; left: 0; right: 0; bottom: 0; }' \
              '</style></head>' \
              '<body>' '<div class="main" style="font-family:arial,sans-serif;' \
              'font-size:13px;line-height:16px;' \
              'color:#222;width:100%;max-width:600px"/>' \
              '</div></body></html>'
    soup = BeautifulSoup(htmlstr, 'html.parser')
    strHd = 'Google Scholar Alert Collection'  # page title
    strRg = 'From ' + trueDateTimeRange[0].strftime("%Y-%m-%d %H:%M:%S") + ' to ' + \
            trueDateTimeRange[1].strftime("%Y-%m-%d %H:%M:%S") # html time range
    # strRg = 'From ' + publications[-1].dateLists[-1].strftime("%Y-%m-%d %H:%M:%S")+ ' to '  +\
    #     publications[0].dateLists[0].strftime("%Y-%m-%d %H:%M:%S")
    # add html title in head
    a = soup.new_tag('title')
    a.insert(0, strHd+'  -  '+strRg)
    soup.html.head.insert(0, a)
    # add date time range
    a = soup.new_tag("div")
    a['style'] = "font-family:arial,sans-serif;text-align: left; background-color: #fff; " \
                 "margin: auto; position: absolute; top: 50px; left: 0; right: 0; bottom: 0;" \
                 "font-size:20px;line-height:40px;color:#222;width:100%;max-width:600px"
    a.insert(0, strRg)
    soup.html.body.insert(0, a)
    # add page title
    a = soup.new_tag("div")
    a['style'] = "font-family:arial,sans-serif;text-align: left;  background-color: #fff;" \
                 " margin: auto; font-weight:bold;position: absolute; top: 00px; left: 0; right: 0; bottom: 0;" \
                 "font-size:30px; line-height:60px; color:#1a0dab;width:100%;max-width:600px"
    a.insert(0, strHd)
    soup.html.body.insert(0, a)
    # saveSoupTag(soup)


    idx_pub = 0
    for i in range(len(sorted_idx)):
        pub = publications[sorted_idx[i]]
        if (min(pub.dateLists).date()-dateRange[0]).days >= 0 and \
                (min(pub.dateLists).date()-dateRange[1]).days <= 0:
            idx_pub += 1
            pubTag = copy.copy(pub.soup.html.body.div)
            a = soup.new_tag('span')
            a['style'] = "font-size:11px;font-weight:bold;color:#1a0dab;vertical-align:2px"
            # a['class'] = "title1"
            strNum = ('%d'%(idx_pub))+'.  '
            a.insert(0, strNum)
            pubTag.h3.insert(0, a)
            soup.html.body.div.next_sibling.next_sibling.append(pubTag)
            for j in range(len(pub.dateLists)):
                a = soup.new_tag("div")
                a['style'] = "font-family:arial,sans-serif;font-size:13px;line-height:18px;color:#993456"
                str = pub.dateLists[j].strftime("%Y-%m-%d, %H:%M:%S")\
                      + ' -- ' + pub.subjects[j]
                a.insert(0, str)
                soup.html.body.div.next_sibling.next_sibling.append(a)
            pubTag = copy.copy(pub.soup.html.body.div.find_next_sibling('div'))
            soup.html.body.div.next_sibling.next_sibling.append(pubTag)
            soup.html.body.div.next_sibling.next_sibling.append(soup.new_tag('br'))
    if fileNameHtml == 0:
        fileNameHtml = ospath.join(file_dir, 'html/'+correct_FileName(strRg,'_')+'.html')
    if not ospath.exists('html/'):
        makedirs('html/')
    saveSoupTag(soup, fileNameHtml)
    return fileNameHtml

# save the souptag file as html for test use
def saveSoupTag(soup, fileNameHtml = 'html/temp.html'):
    HTML_str = soup.prettify()
    with open(fileNameHtml, 'w', encoding='utf-8') as f:
        f.write(HTML_str)

# Get the ids from Gmail API according the lables
def GetLabelsId(service, user_id, label_names=[]):
    results = service.users().labels().list(userId=user_id).execute()
    labels = results.get('labels', [])

    label_ids = []
    for name in label_names:
        for label in labels:
            if label['name'] == name:
                label_ids.append(label['id'])
    return label_ids

# join the messages into the stored messages
def joinMsgs(messages, messagesOld):
    for i in range(len(messages)):
        if messages[-1]['id'] == messagesOld[i]['id']:
            break
    messagesJoin = messages+messagesOld[i+1:]
    # messagesJoin = messages+messagesOld[i:-1]
    # print(messagesJoin[i-1:i+3])
    # print(messagesOld[i-1:i+3])
    return messagesJoin

# get all the messagesID labeled as unread and send from scholar email
def ListMessagesWithLabels(service, user_id, labels, messagesOld):
    """List all Messages of the user's mailbox with label_ids applied.
    Args:
      service: Authorized Gmail API service instance.
      user_id: User's email address. The special value "me"
      can be used to indicate the authenticated user.
      label_ids: Only return Messages with these labelIds applied.
    Returns:
      List of Messages that have all required Labels applied. Note that the
      returned list contains Message IDs, you must use get with the
      appropriate id to get the details of a Message.
    """
    messages = list()
    msgDictOld = mkMsgDict(messagesOld)
    label_ids = GetLabelsId(gmail, user_id, labels)
    try:
        response = service.users().messages().list(userId=user_id,
                                                   labelIds=label_ids,
                                                   q="from:" + scholar_email).execute()
        if 'messages' in response:
            messages.extend(response['messages'])
        while 'nextPageToken' in response:
            page_token = response['nextPageToken']
            response = service.users().messages().list(userId=user_id,
                                                       labelIds=label_ids,
                                                       pageToken=page_token).execute()
            messages.extend(response['messages'])
            if math.fmod(len(messages), 200) == 0:
                print('Got %d messages ID' % len(messages))
            if messages[-1]['id'] in msgDictOld:
                messages = joinMsgs(messages, messagesOld)
                break
        return messages
    except Exception as error:
        print('An error occurred: %s' % error)

# make the dict of ids of the messages
def mkMsgDict(messages):
    MsgDict = dict()
    if messages is None:
        return MsgDict
    for i in range(len(messages)):
        messages[i]['id']
        MsgDict[messages[i]['id']] = i
    return MsgDict

# make the dict of titles of the publications
def mkPubTitDict(publications):
    pubTitDict = dict()
    if publications is None:
        return pubTitDict
    for i in range(len(publications)):
        pubTitDict[publications[i].bib['title']] = i
    return pubTitDict


# make the dict of ids of the publications
def mkPubIdDict(publications):
    pubIdDict = dict()
    if publications is None:
        return pubIdDict
    for i in range(len(publications)):
        for id in publications.messageIDs:
            pubIdDict[id] = i
    return pubIdDict


# Save the messages, scholarMessages, publications to pkl file,
# if no cache pkl exists, create a new file to pkl folder
def pklLoad(pklFileName):
    if os.access(pklFileName, os.F_OK):
        with open(pklFileName, 'rb') as f:
            pkls = pickle.load(f)
            messages = pkls[0]
            scholarMessages = pkls[1]
            publications = pkls[2]
            # chack the type of the loaded, clear if not a list
            if type(messages) != type([]):
                message = list()
            if type(scholarMessages) != type([]):
                scholarMessages = list()
            if type(publications) != type([]):
                publications = list()
            # sglMsgDict = mkMsgDict(scholarMessages)
            # pubTitDict = mkPubTitDict(publications)
            # return messages, scholarMessages, publications, sglMsgDict, pubTitDict
            return messages, scholarMessages, publications
    else:
        if not ospath.exists('pkl/'):
            makedirs('pkl/')
        messages = list()
        scholarMessages = list()
        publications = list()
        # sglMsgDict = dict()
        # pubTitDict = dict()
        # return messages, scholarMessages, publications, sglMsgDict, pubTitDict
        return messages, scholarMessages, publications

# save messages, scholarMessages, publications as pkl into pkl folder
def pklSave(pklFileName, messages, scholarMessages, publications):
    scholarMessages = sorted(scholarMessages, key=lambda k: k['id'], reverse=True)
    pkls = (messages, scholarMessages, publications)
    with open(pklFileName, 'wb') as f:
        pickle.dump(pkls, f)

# Use regular expressions to get the Alert author and Alert type
# of the article email subject lists
def pubSub2AuthorType(subListStr):
    AuthorTypeList = list()
    # re_name = '( ?\w+)( ?\w+)(-? ?\w+)'
    re_name = '[\w\' .-]+'
    re_ctbyEn = '(?<=to articles by )'+re_name
    re_ctbyZh = re_name+'(?=的文章新增了 )'
    re_nAtcEn = re_name+'(?= - new article)'
    re_nAtcZh = re_name+'(?= - 新文章)'
    re_rRshEn = re_name+'(?= - new related research)'
    re_rRshZh = re_name+'(?= - 新的相关研究工作)'
    for i in range(len(subListStr)):
        auth = ''
        type = ''
        resultCE = re.search(re_ctbyEn, subListStr[i])
        resultCZ = re.search(re_ctbyZh, subListStr[i])
        resultNE = re.search(re_nAtcEn, subListStr[i])
        resultNZ = re.search(re_nAtcZh, subListStr[i])
        resultRE = re.search(re_rRshEn, subListStr[i])
        resultRZ = re.search(re_rRshZh, subListStr[i])
        if resultCE:
            auth = resultCE[0]
            type = subType_AuthCitation
        elif resultCZ:
            auth = resultCZ[0]
            type = subType_AuthCitation
        elif resultNE:
            auth = resultNE[0]
            type = subType_AuthNew
        elif resultNZ:
            auth = resultNZ[0]
            type = subType_AuthNew
        elif resultRE:
            auth = resultRE[0]
            type = subType_AuthRelated
        elif resultRZ:
            auth = resultRZ[0]
            type = subType_AuthRelated
        AuthorTypeList.append([auth, type])
    return AuthorTypeList

# change a list of strings to a list of list string
def listOfList(strList = list()):
    listList = list()
    for str in strList:
        listList.append([str])
    return listList

# save the CSV file to csv folder with header in list of str
# and columns of data list as list of str or list of list of str
def saveCSV(fileName, header, list1=list(), list2=list()):
    data = list()
    if type(list1[0]) == type(str()):
        list1 = listOfList(list1)
    if type(list2[0]) == type(str()):
        list2 = listOfList(list2)
    for i in range(len(list1)):
        data.append(list1[i]+list2[i])
    with open(fileName, 'w', newline='', encoding='gb18030') as csvfile:  # gb2312 gb18030 utf-8
        f_csv = csv.DictWriter(csvfile, header)
        f_csv.writeheader()
        spamwriter = csv.writer(csvfile)
        spamwriter.writerows(data)

# get the author-value list from AuthVal.csv in csv folder,
# if AuthVal.csv is not exist, create a new AuthVal.csv
# using the alert subjects in publications
def getPubATcsv(publications):
    authVal = dict()
    if os.access('csv/AuthVal.csv', os.F_OK):
        authValHd = list()
        authValList = list()
        with open('csv/AuthVal.csv', encoding='gb18030') as f:
            f_csv = csv.reader(f)
            authValHd = next(f_csv)
            for row in f_csv:
                authValList.append(row)
        authValList = sorted(authValList,
                             key=lambda k: k[1].split(' ')[-1])
        for avl in authValList:
            authVal[avl[1]] = float(avl[0])
    else:
        if not ospath.exists('csv/'):
            makedirs('csv/')
        subDict = dict()
        for pub in publications:
            for sub in pub.subjects:
                subDict[sub] = 1
        subListStr = [i for i in subDict.keys()]
        AuthorTypeList = pubSub2AuthorType(subListStr)
        # saveCSV('csv/Auth_Type_Sub.csv', ['Author', 'Type', 'Subject'], AuthorTypeList, subListStr)
        authDict = dict()
        for auth in AuthorTypeList:
            authDict[auth[0]] = 1
        authList = [i for i in authDict.keys()]
        authList = sorted(authList,
                             key=lambda k: k.split(' ')[-1])
        authValList = ['1.0']*len(authList)
        if not ospath.exists('csv/'):
            makedirs('csv/')
        saveCSV('csv/AuthVal.csv', ['Value', 'Author'], authValList, authList)
        for i in range(len(authList)):
            authVal[authList[i]] = 1.0
    return authVal



# get the Journal-Value list from JonlVal.csv in csv folder,
# if AuthVal.csv is not exist, create a new AuthVal.csv
# using the alert subjects in publications
def getPubJonlcsv(publications):
    jonlVal = dict()
    if os.access('csv/JonlVal.csv', os.F_OK):
        jonlValHd = list()
        jonlValList = list()
        with open('csv/JonlVal.csv', encoding='gb18030') as f:
            f_csv = csv.reader(f)
            jonlValHd = next(f_csv)
            for row in f_csv:
                jonlValList.append(row)
        jonlValList = sorted(jonlValList,
                             key=lambda k: k[1])
        for avl in jonlValList:
            jonlVal[avl[1]] = float(avl[0])
    else:
        if not ospath.exists('csv/'):
            makedirs('csv/')
        jonlDict = dict()
        for pub in publications:
            if 'journal' in pub.bib:
                jonlDict[pub.bib['journal']] = 1
        jonlListStr = [i for i in jonlDict.keys()]
        jonlListStr = sorted(jonlListStr)
        jonlValList = ['1.0']*len(jonlListStr)
        if not ospath.exists('csv/'):
            makedirs('csv/')
        saveCSV('csv/JonlVal.csv', ['Value', 'Journal'], jonlValList, jonlListStr)
        for i in range(len(jonlListStr)):
            jonlVal[jonlListStr[i]] = 1.0
    return jonlVal


# # sort the val_string list
# def sorted_valStr(valStrList):
#     sorted_idx = sorted(range(len(valStrList)),
#                        key=lambda k: valStrList[k][1])
#     sorted_valStrList = sorted(valStrList, key=lambda k: k[1].split(' ')[-1], reverse=False)
#     return [sorted_valStrList, sorted_idx]



if __name__ == '__main__':
    # load all the cached data

    pklFileName = 'pkl/pkl_data.pkl'
    messages, scholarMessages, publications = pklLoad(pklFileName)

    # get the GmailApi service
    gmail = getGmailApi()

    # Get all the messages with CATEGORY_UPDATES labels
    print('List messagesID using gmail API')
    t.tic()
    messages = ListMessagesWithLabels(gmail, user_id, ['CATEGORY_UPDATES'], messagesOld=messages)
    t.toc()
    pklSave(pklFileName, messages, scholarMessages, publications)
    print('Found %d messages' % len(messages))

    # pull the scholarMessages from GmailApi
    maxDays = 65
    maxRange = 5000
    scholarMessages = pullMessage(gmail, messages, maxDays, maxRange, scholarMessages)
    pklSave(pklFileName, messages, scholarMessages, publications)
    print('Pull %d messages' % len(scholarMessages))

    # Parse the scholarMessages get Publications and pubTitDict
    t.tic()
    forceRead = False
    # forceRead = True
    publications = msg2Pub(scholarMessages, publications, forceRead)[0]  #[publications, pubTitDict, nMsgParsed]
    t.toc()
    pklSave(pklFileName,messages, scholarMessages, publications)
    print('Got %d Pubs' % len(publications))

    # rating the Pubs
    authVal = getPubATcsv(publications)
    jonlVal = getPubJonlcsv(publications)
    t.tic()
    [scorePubs, sorted_scorePubs, sorted_idx] = rateSortPubs(publications, authVal, jonlVal)
    t.toc()
    print('Sorte the Pubs')
    # print(str(sorted_scorePubs) + '\n' + str(sorted_idx))

    # save the html file using the dateRange
    # fileNameHtml = 'html/html_soup_joint1.html'
    dateRange = [date(2021,4,1), date(2021,4,15)]
    # dateRange = [date.today() - timedelta(days=2), date.today()]
    t.tic()
    fileNameHtml = savPub2html(publications, sorted_idx,0, dateRange)
    t.toc()
    print('Html file saved at %s' %fileNameHtml)


    web = webbrowser.get(BROWSER_COMMAND)
    try:
        web.open(fileNameHtml, new = 2)
    except:
        print('Open the html file failed')

    # # Test of Rate score
    # idx = 108
    # publications[idx].subjects
    # publications[idx].ratingScore(authVal)
    # for pub in publications:
    #     pub.ratingScore(authVal)
    #     print(pub.score)

    # # TO DO : encode of Fernández
    # re_name = '[\w\' .-]+'
    # re_ctbyZh = re_name + '(?=的文章新增了 )'
    # a = 'María R. Fernández-Ruiz的文章新增了 2 次引用'
    # a2 = 'Regina Magalhães的文章新增了 1 次引用'
    # result = re.match(re_ctbyZh, a2)
    # if result:
    #     print(result[0])
    # a1 = list()
    # a2 = list()
    # with open('eggs.csv',encoding='gb18030') as f:
    #     f_csv = csv.reader(f)
    #     a1 = next(f_csv)
    #     for row in f_csv:
    #         a2.append(row)