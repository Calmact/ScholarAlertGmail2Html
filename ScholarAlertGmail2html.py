from __future__ import print_function
import numpy as np
import pickle
import pandas as pd
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from apiclient import errors
from bs4 import BeautifulSoup
# from bs4 import Tag
import base64
import urllib
from retrying import retry
import math
import re
from datetime import datetime, date, timedelta
from pytz import timezone
import csv
import copy
from pytictoc import TicToc
import webbrowser
# from matplotlib import pyplot as plt

# # Use http proxy as a global proxy
# os.environ["http_proxy"] = "http://127.0.0.1:10809"
# os.environ["https_proxy"] = "http://127.0.0.1:10809"

file_dir = os.path.dirname(__file__)
SCOPES = ['https://www.googleapis.com/auth/gmail.modify'] # gmail.modify  gmail.readonly
user_id = 'me'
scholar_email = 'scholaralerts-noreply@google.com'
cst_tz = timezone('Asia/Shanghai')
BROWSER_COMMAND = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
t = TicToc()

subType_AuthCitation = 'citation'
subType_AuthNew = 'new article'
subType_AuthRelated = 'related'
subType_ArticleCitations = 'new citations'
subType_kewWordNewResults = 'new results'
typeVal = dict()
typeVal[subType_AuthNew] = 8
typeVal[subType_AuthCitation] = 2
typeVal[subType_AuthRelated] = 0.5
typeVal[subType_ArticleCitations] = 6
typeVal[subType_kewWordNewResults] = 1.5
typeVal[''] = 1

class sag2hException(Exception):
    # print(Exception)
    pass

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


# Mark email as being read (remove unread label)
def markRead(gmail, message_description):
    message_id = message_description['id']
    body = {'removeLabelIds' : ['UNREAD']}  #  body = {"addLabelIds": [], "removeLabelIds": ["UNREAD", "INBOX"]}
    return gmail.users().messages().modify(id=message_id, userId=user_id, body=body).execute()

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

# Get the service interface of GmailApi, validate and save
# the token.json and credentials.json files of GmailAPI
# retry https://www.biaodianfu.com/python-error-retry.html
@retry(wait_random_min=100, wait_random_max=1000, stop_max_attempt_number=4)
def getGmailApi():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if not os.path.exists('json/'):
        os.makedirs('json/')
    if os.path.exists('json/token.json'):
        creds = Credentials.from_authorized_user_file('json/token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if os.access('json/credentials.json', os.F_OK):
                flow = InstalledAppFlow.from_client_secrets_file(
                    'json/credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            else:
                print('No credentials.json file, please download your credentials.json to the json folder')
                raise sag2hException(1)
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
        urlstrip = re.findall('(?<=url=).*?(?=&)', urlstr)
        if len(urlstrip) > 0:
            self.bib['url'] = urllib.parse.unquote(urlstrip[0])
        else:
            self.bib['url'] = urlstr

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
        self.authTypeList = []
        self.typeScores = []
        self.authScores = []
        self.subjectScore = 0.0
        self.score = 0.0
        self.jonlScore = 0.0

        # add a soup
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
        # TODO_done delete the br tag in the abstract
        for i in abstractTag.find_all('br'):
            i.decompose()
        self.soup.html.body.div.append(abstractTag)
        self.soup.html.body.insert(1, shareTag)
        # [基础]-beautifulsoup模块使用详解 https://blog.csdn.net/lu8000/article/details/82313054
        # saveSoupTag(self.soup)
        # saveSoupTag(abstractTag)
        # a = self.soup.prettify()

    def addHeaders(self, sMsg):
        headers = sMsg['payload']['headers']
        subject = getSubject(headers)
        datetimeMsg = getDate(headers)
        self.subjects.append(subject)
        self.messageIDs.append(sMsg['id'])
        self.dateLists.append(datetimeMsg)

    def ratingSubJonl(self, authVal, jonlVal):
        self.score = 0.0
        self.jonlScore = 0.0
        self.subjectScore = 0.0
        self.typeScores = []
        self.authScores = []
        self.authTypeList = pubSub2AuthorType(self.subjects)
        for i in range(len(self.subjects)):
            # [authScore, typeScore] = rateSub(self.subjects[i], authVal)
            # self.typeScores.append(typeScore)
            # self.authScores.append(authScore)
            self.authScores.append(authVal[self.authTypeList[i][0]])
            self.typeScores.append(typeVal[self.authTypeList[i][1]])
            self.subjectScore += self.authScores[i] * self.typeScores[i]
        self.jonlScore = rateJonl(self, jonlVal)

    def ratingScore(self, authVal, jonlVal, k_factor):
        authDict = dict()
        for authType in self.authTypeList:
            authDict[authType[0]] = 1.0
        self.ratingSubJonl(authVal, jonlVal)
        self.score = self.subjectScore*len(authDict)*self.jonlScore
        return self.score

# Use regular expressions to get the Alert author and Alert type
# of the article email subject lists
def pubSub2AuthorType(subjectsList):
    AuthorTypeList = list()
    # re_name = '( ?\w+)( ?\w+)(-? ?\w+)'
    re_name = '\w+[\w\' .-]+'
    # re_ctbyEn = '(?<=to articles by )'+re_name
    # re_ctbyZh = re_name+'(?=的文章新增了 )'
    # re_nAtcEn = re_name+'(?= - new article)'
    # re_nAtcZh = re_name+'(?= - 新文章)'
    # re_rRshEn = re_name+'(?= - new related research)'
    # re_rRshZh = re_name+'(?= - 新的相关研究工作)'
    #
    # re_nCtEn = '(?=- new citations)'
    # re_nCtZh = '(?=- new citations)'
    # re_nRtEn = '(?=- new results)'
    # re_nRtZh = '(?= - 新的结果)'
    for i in range(len(subjectsList)):
        auth = ''
        type = ''
        resultCE = re.search('(?<=to articles by )'+re_name, subjectsList[i])
        resultCZ = re.search(re_name+'(?=的文章新增了 )', subjectsList[i])
        resultNE = re.search(re_name+'(?= - new article)', subjectsList[i])
        resultNZ = re.search(re_name+'(?= - 新文章)', subjectsList[i])
        resultRE = re.search(re_name+'(?= - new related research)', subjectsList[i])
        resultRZ = re.search(re_name+'(?= - 新的相关研究工作)', subjectsList[i])

        resultnCtE = re.search('(?=- new citations)', subjectsList[i])
        resultnCtZ = re.search('(?=- 新的引用)', subjectsList[i])
        resultnRtE = re.search('(?=- new results)', subjectsList[i])
        resultnRtZ = re.search('(?= - 新的结果)', subjectsList[i])
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

        elif resultnCtE or resultnCtZ:
            auth = 'Articles'
            type = subType_ArticleCitations
        elif resultnRtE or resultnRtZ:
            auth = 'Key words'
            type = subType_kewWordNewResults
        AuthorTypeList.append([auth, type])
    return AuthorTypeList

# parse a single subject and get the corresponding [authScore, typeScore].
def rateSub(subject, authVal):
    authScore = 1.0
    typeScore = 1.0
    [[auth, type]] = pubSub2AuthorType([subject])
    if auth in authVal:
        authScore = authVal[auth]
    if type in typeVal:
        typeScore = typeVal[type]
    return [authScore, typeScore]

# parse a single publication, get the jonlScore
def rateJonl(publication, jonlVal):
    jonlScore = 1.0
    if 'journal' in publication.bib:
        if publication.bib['journal'] in jonlVal:
            jonlScore = jonlVal[publication.bib['journal']]
    return jonlScore

# rate the publications and get the soted scores and the sorted idxe
def rateSortPubs(publications, authVal, jonlVal):
    for idx in range(len(publications)):
        publications[idx].ratingSubJonl(authVal, jonlVal)
    k_factor = scoreFactor(publications)
    scorePubs = [0.0 for i in range(len(publications))]
    for idx in range(len(publications)):
        scorePubs[idx] = publications[idx].ratingScore(authVal, jonlVal, k_factor)
    sorted_idx = sorted(range(len(scorePubs)),
                       key=lambda k: publications[k].score,
                       reverse=True)
    sorted_scorePubs = sorted(scorePubs, reverse=True)
    return [scorePubs, sorted_scorePubs, sorted_idx]

# save the pub to html file according to date range
def savPub2html(publications, sorted_idx, fileNameHtml = 0, dateRange=0, authVal = dict(), jonlVal = dict()):
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
            pubDatetimeMax.append(min(pub.dateLists))
    trueDateTimeRange[0] = min(pubDatetimeMin)
    trueDateTimeRange[1] = max(pubDatetimeMax)
    # html head
    htmlstr = ' <html xmlns="http://www.w3.org/1999/xhtml" xmlns:o="urn:schemas-microsoft-com:office:office">' \
              '<head> <meta http-equiv="Content-Type" content="text/html;charset=UTF-8" /> <style>' \
              'gse_alrt_sni{text-align:justify}' \
              '.main{ text-align:justify;  background-color: #fff; margin: auto; ' \
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
        if pub.score == 0 or pub.subjectScore == 0 or pub.jonlScore == 0:
            continue
        if (min(pub.dateLists).date()-dateRange[0]).days >= 0 and \
                (min(pub.dateLists).date()-dateRange[1]).days <= 0:
            # copy the pub tag from the email
            pubTag = copy.copy(pub.soup.html.body.div)
            # insert a idx before the title
            idx_pub += 1
            a = soup.new_tag('span')
            a['style'] = "font-size:11px;font-weight:bold;color:#1a0dab;vertical-align:2px"
            strNum = ('%d'%(idx_pub))+'.  '
            a.insert(0, strNum)
            pubTag.h3.insert(0, a)

            # add the title/author/abstract Tag
            pubTag.div.next_sibling['style'] = 'text-align:justify'
            pubTag.h3.a['href'] = pub.bib['url']
            # add the subject and the datetime of the subject
            sort_subjectScore_idx = sorted(range(len(pub.dateLists)),
                                    key=lambda k:
                                    [pub.authScores[k]*pub.typeScores[k],
                                    pub.typeScores[k],
                                    pub.authScores[k]],
                                    reverse=True)
            subScore = 0
            for j in sort_subjectScore_idx:
                a = soup.new_tag("div")
                a['style'] = "font-family:arial,sans-serif;font-size:13px;line-height:18px;color:#993456"
                # TODO_done add the subject scores (authScore and typeScore)
                # [[auth, type]] = pubSub2AuthorType([pub.subjects[j]])
                [auth, type] = pub.authTypeList[j]
                str_add = pub.dateLists[j].strftime("%Y-%m-%d, %H:%M:%S")\
                      +' -- '+ "{:3.1f}".format(authVal[auth]*typeVal[type])\
                      + ' -- ['+ auth +' | '+ type + '] '\
                      +'[' +"{:02.1f}".format(authVal[auth])+ ' | '+"{:02.1f}".format(typeVal[type])\
                      +']' + ' -- '+ pub.subjects[j]
                a.insert(0, str_add)
                pubTag.append(a)
                subScore += authVal[auth]*typeVal[type];
            a = soup.new_tag("div")
            a['style'] = "font-family:arial,sans-serif;" \
                         "font-size:13px;line-height:18px;color:#993456"
            str_add = "<div><b>"+"{:3.1f}".format(pub.score)+'</b><sub>(Totle Score)</sub>; '+\
                      "<b>"+"{:3.1f}".format(subScore)+'</b><sub>((Subject Score)</sub>; '+\
                      "<b>"+"{:3.1f}".format(pub.jonlScore)+'</b><sub>(Journal Score)</sub></div>' #×
            a.append(BeautifulSoup(str_add, 'html.parser'))
            # saveSoupTag(a)
            pubTag.h3.next_sibling.next_sibling.append(a)
            pubTag.append(copy.copy(pub.soup.html.body.div.find_next_sibling('div')))
            pubTag.append(soup.new_tag('br'))
            # saveSoupTag(pubTag)
            # saveSoupTag(soup)
            soup.html.body.div.next_sibling.next_sibling.append(pubTag)
            # soup.html.body.div.next_sibling.next_sibling.append(soup.new_tag('br'))
    if fileNameHtml == 0:
        fileNameHtml = os.path.join(file_dir, 'html/'+correct_FileName(strRg,'_')+'.html')
    if not os.path.exists('html/'):
        os.makedirs('html/')
    saveSoupTag(soup, fileNameHtml)
    return fileNameHtml

# save the souptag file as html for Logs use
def saveSoupTag(soup, fileNameHtml = 'html/temp.html'):
    HTML_str = soup.prettify()
    with open(fileNameHtml, 'w', encoding='utf-8') as f:
        f.write(HTML_str)

# Validate the gmail labels
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
    # query="from:" + scholar_email
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
            # check the type of the loaded, clear if not a list
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
        if not os.path.exists('pkl/'):
            os.makedirs('pkl/')
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
    if len(list1)>0 and type(list1[0]) == type(str()):
        list1 = listOfList(list1)
    if len(list2)>0 and type(list2[0]) == type(str()):
        list2 = listOfList(list2)
    for i in range(len(list1)):
        data.append(list1[i]+list2[i])
    with open(fileName, 'w', newline='', encoding='utf-8_sig') as csvfile:  # gb2312 gb18030 utf-8
        f_csv = csv.DictWriter(csvfile, header)
        f_csv.writeheader()
        spamwriter = csv.writer(csvfile)
        spamwriter.writerows(data)

# get the author/journal-value list from AuthVal/JonlVal.csv in csv folder,
# if AuthVal/Jonl.csv is not exist, create a new AuthVal/JonlVal.csv
# using the alert subjects in publications for author in AuthVal
def getAuthJonlcsv(publications, filePathName ='csv/AuthVal.csv'):
    fileName = filePathName.split('.')[0].split('/')[1]
    # TODO pandas the csvs
    # pdcsv = pd.read_csv(filePathName)
    ajValDict = dict()
    if os.access(filePathName, os.F_OK):
        ajValHd = list()
        ajValList = list()
        with open(filePathName, encoding='utf-8_sig') as f:  #encoding='gb18030'
            f_csv = csv.reader(f)
            ajValHd = next(f_csv)
            for row in f_csv:
                ajValList.append(row)
        for i in range(len(ajValList)-1, -1, -1):
            if len(ajValList[i]) != 2:
                ajValList.pop(i)
        if 'Auth' in filePathName:
            ajValList = sorted(ajValList,
                                 key=lambda k: k[1].split(' ')[-1])
        elif 'Jonl' in filePathName:
            ajValList = sorted(ajValList,
                                 key=lambda k: k[1])
        for avl in ajValList:
            ajValDict[avl[1]] = float(avl[0])
    else:
        if not os.path.exists('csv/'):
            os.makedirs('csv/')
        ajDict = ajDictInit(publications, filePathName)
        ajList = [i for i in ajDict.keys()]
        for i in range(len(ajList)):
            ajValDict[ajList[i]] = 1.0
        saveAjvDict(ajValDict, filePathName)
    return ajValDict

def ajDictInit(publications, filePathName):
    ajDict = dict()
    if 'Auth' in filePathName:
        subDict = dict()
        for pub in publications:
            for sub in pub.subjects:
                subDict[sub] = 1
        subListStr = [i for i in subDict.keys()]
        AuthorTypeList = pubSub2AuthorType(subListStr)
        # saveCSV('csv/Auth_Type_Sub.csv', ['Author', 'Type', 'Subject'], AuthorTypeList, subListStr)
        for auth in AuthorTypeList:
            ajDict[auth[0]] = 1.0
    elif 'Jonl' in filePathName:
        for pub in publications:
            if 'journal' in pub.bib:
                ajDict[pub.bib['journal']] = 1.0
    return ajDict


def loadAuthJonlVal(publications, fileName ='AuthVal'):
    ajValDict = getAuthJonlcsv(publications,'csv/'+fileName+'.csv')
    ajDict = ajDictInit(publications, 'csv/' + fileName + '.csv')
    for aj in ajDict:
        if aj not in ajValDict:
            ajValDict[aj] = 1.0
    if os.access('csv/' + fileName + '-simplify.csv', os.F_OK):
        ajVal_simpl = getAuthJonlcsv(publications,'csv/'+fileName+'-simplify.csv')
        #update condition.
        # force update
        for ajVo in ajVal_simpl:
            ajValDict[ajVo] = ajVal_simpl[ajVo]
        # save the ajValDict and simplified ajValDict
        saveAjvDict(ajValDict, 'csv/' + fileName + '.csv')
        ajValDict_simpl = ajValDict.copy()
        for key in list(ajValDict_simpl):
            if ajValDict_simpl[key] == 1:
                ajValDict_simpl.pop(key)
        saveAjvDict(ajValDict_simpl, 'csv/' + fileName + '-simplify.csv')
    if os.access('csv/'+fileName+'-backup.csv', os.F_OK):
        ajVal_old = getAuthJonlcsv(publications,'csv/'+fileName+'-backup.csv')
        #update condition.
        # a. key exists in ajValDict. force update if the ajValDict key value is 1.
        # b. New key value does not exist: create New key and value
        for ajVo in ajVal_old:
            if ajVo in ajValDict:
                if ajValDict[ajVo] == 1:
                    ajValDict[ajVo] = ajVal_old[ajVo]
            else:
                ajValDict[ajVo] = ajVal_old[ajVo]
    # save the ajValDict and simplified ajValDict
    saveAjvDict(ajValDict, 'csv/' + fileName + '.csv')
    ajValDict_simpl = ajValDict.copy()
    for key in list(ajValDict_simpl):
        if ajValDict_simpl[key] == 1:
            ajValDict_simpl.pop(key)
    saveAjvDict(ajValDict_simpl, 'csv/' + fileName + '-simplify.csv')
    return ajValDict

def saveAjvDict(ajValDict, filePathName):
    ajList = [i for i in ajValDict.keys()]
    if 'Auth' in filePathName:
        if 'simp' in filePathName:
            ajList = sorted(ajList,
                            key=lambda k: ajValDict[k], reverse=True)
        else:
            ajList = sorted(ajList,
                            key=lambda k: k.split(' ')[-1])
        valList = [str(ajValDict[i]) for i in ajList]
        saveCSV(filePathName, ['Value', 'Author'], valList, ajList)
    elif 'Jonl' in filePathName:
        if 'simp' in filePathName:
            ajList = sorted(ajList,
                            key=lambda k: [ajValDict[k],ajList], reverse=True)
        else:
            ajList = sorted(ajList)
        valList = [str(ajValDict[i]) for i in ajList]
        saveCSV(filePathName, ['Value', 'Jonl'], valList, ajList)

def scoreFactor(publications):
    idx_np = np.arange(0, len(publications))
    scoreArray = np.zeros([len(publications), 3])
    for i in range(len(publications)):
        scoreArray[i, 0] = publications[i].subjectScore
        scoreArray[i, 1] = publications[i].jonlScore
        # scoreArray[i, 2] = publications[i].subjects
    a = -np.sort(-scoreArray[:, 0:2].T).T
    a20 = a[0:int(np.floor(a[:,0].size / 5)),:]
    a.mean(axis=0)
    a.std(axis=0)
    a20.mean(axis=0)
    a20.std(axis=0)
    # plt.plot(idx_np/idx_np.max()*100,a)
    # plt.xlabel('Index pencentage (%)')
    # plt.ylabel('Subject score (a.u.)')
    # plt.title('_Min '+str(a20.min(axis = 0)[0])+'  Max '+str(a20.max(axis = 0)[0])+'/'+
    #           ' Min '+str(a20.min(axis = 0)[1])+'  Max '+str(a20.max(axis = 0)[1]))
    # plt.xlim([0, 20])
    # plt.grid()
    # plt.show()

    k_factor =  a20.mean(axis=0)[0] / a20.mean(axis=0)[1]
    scoreArray[:, 2] = scoreArray[:, 0] + scoreArray[:, 1] * k_factor
    sortIdx = np.argsort(-scoreArray[:, 2])
    # plt.plot(idx_np / idx_np.max() * 100, scoreArray[sortIdx,:])
    # plt.grid()
    # plt.show()
    return float(k_factor)


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
    # dateRange = [date(2021,1,1), date(2021,6,1)]
    dateRange = [date.today() - timedelta(days=1), date.today()]
    maxDays = (date.today()-dateRange[0]).days
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

    # # rating the Pubs
    authVal = loadAuthJonlVal(publications, 'AuthVal')
    jonlVal = loadAuthJonlVal(publications, 'JonlVal')
    t.tic()
    [scorePubs, sorted_scorePubs, sorted_idx] = rateSortPubs(publications, authVal, jonlVal)
    t.toc()
    print('Sorte the Pubs')
    # print(str(sorted_scorePubs) + '\n' + str(sorted_idx))

    # save the html file using the dateRange
    # fileNameHtml = 'html/html_soup_joint1.html'
    t.tic()
    fileNameHtml = savPub2html(publications, sorted_idx, 0, dateRange, authVal, jonlVal)
    t.toc()
    print('Html file saved at %s' %fileNameHtml)

    web = webbrowser.get(BROWSER_COMMAND)
    try:
        web.open(fileNameHtml, new = 2)
    except:
        print('Open the html file failed')

    # for message in messages[0:10]:
    #     markRead(gmail, message)