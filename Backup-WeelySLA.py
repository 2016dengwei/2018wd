# -*- coding: utf-8 -*-
"""
Created on  Mar 22 2018

@author: Wei Deng

weekly SLA list
backup version, with no username and pwd
"""



#import all need lib here

import pandas as pd
import atws
import numpy as np
import datetime
import pytz
import csv





#def getReportTicketList(reportYear,reportMonth, lastDay):
def getSLATicketList():
    #get report month( last month) WHOLE ticket list from autotast

    #create new empty dataframe
    #FRDateTime: first response date and time
    #SLVDateTime: solved date and time
    
    col=['TKNumber', 'AccountName', 'Priority','Status','ContactName','Source','Issue','SubIssue','FRDateTime','FRDueDateTime','FR_SLAMet','SLVDateTime,','SLVDueDateTime','Actual SLA Met Tickets','QueueID','Final']
    temp=pd.DataFrame(columns=col)
    
    start,end, filedate=getCreateDateTime()
    
    startDate=getEstDateTime(start)
    endDate=getEstDateTime(end)
    #startDate=str(reportYear)+'-'+reportMonth+'01'
    #endDate=str(reportYear)+'-'+reportMonth+str(lastDay)
    
    at=getConnect()

#use resolvedDateTime<REsolvedDueDateTime check met SLA or not
#use FirstResponseDateTime<FirstResponseDueDateTime check met SLA OR NOT

#get ticket list for report month
    tklist=getTicketListByDate(at,startDate,endDate)

#get picklist for name of issue, subissue, status, priority, ticketsource
    plist=getPicklist(at,'Ticket')
    
#get account list from autotask
    accountlist=getAutotaskName(at,'Account')

#get contact list from autotask
    contactlist=getAutotaskName(at,'Contact')
    
    i=0
    for tk in tklist:
#frmet/slvmet=0: sla not met, =1: sla met
        frmet=0
        slvmet=0
        QueueID='NoQueueID'
        
        
        
        issue=plist['IssueType'].reverse_lookup(str(tk['IssueType']))
        try:
            subIssue=plist['SubIssueType'].reverse_lookup(str(tk['SubIssueType']))
        except:
            subIssue='NoSubIssue'
        plevel=plist['Priority'].reverse_lookup(str(tk['Priority']))
        status=plist['Status'].reverse_lookup(str(tk['Status']))
        source=plist['Source'].reverse_lookup(str(tk['Source']))
        
        
        rlvdatetime='no record'
        rlvdue='no record'
        if ('QueueID' in tk):
            QueueID=tk['QueueID']
        
        #if ('ResolvedDateTime' in tk and 'ResolvedDueDateTime' in tk):
        if ('ResolvedDueDateTime' in tk):
            rlvdue=tk['ResolvedDueDateTime']
            
            if ('ResolvedDateTime' in tk): 
                rlvdatetime=tk['ResolvedDateTime']
       
                if tk['ResolvedDateTime']<tk['ResolvedDueDateTime']:
                    slvmet=1
        
        frdatetime='no record'
        frdue='no record'
        if ('FirstResponseDateTime' in tk and 'FirstResponseDueDateTime' in tk):
            frdatetime=tk['FirstResponseDateTime']
            frdue=tk['FirstResponseDueDateTime']
        
            if tk['FirstResponseDateTime']<tk['FirstResponseDueDateTime']:
                frmet=1
        
#get contact name, if no contact, 'empty'
        contactname='empty'
        if ('ContactID' in tk):
            contactname=getContactName(tk['ContactID'], contactlist)
#get accountname from autotask
        accountname='empty'
        if ('AccountID' in tk):
            accountname=getAccountNameByID(str(tk['AccountID']),accountlist)

#get ticket number
        tknum=getTicketNumberByID(at,tk['id'])
        

        temp.loc[str(i)]=[str(tknum),str(accountname),str(plevel), str(status),str(contactname),str(source),str(issue),str(subIssue),str(frdatetime),str(frdue),int(frmet),str(rlvdatetime),str(rlvdue), int(slvmet), str(QueueID),int(frmet)+int(slvmet) ]
        i=i+1
    
    #only return service desk queue , queueid=29683528
    tlist_SD=temp[(temp['QueueID']=='29683528')]
    tklist_NoMonitoring=tlist_SD[(tlist_SD['Issue']!='Monitoring')]
    tk_NoAlert=tklist_NoMonitoring[(tklist_NoMonitoring['SubIssue']!='Solarwinds Alert')]
    tlist=tk_NoAlert[(tk_NoAlert['Source']!='Monitoring Alert')]
    
    filepath="./weekly_workedhours/"+filedate+"_WeeklySLA.csv"
    tlist.to_csv(filepath,index=False)
    
    tk_NoMet=tlist[(tlist['Final']<2)]
    
    filepath="./weekly_workedhours/"+filedate+"_WeeklySLA_NoMet.csv"
    
    #write sla not met data to file used by weekly powerbi file
    tlist.to_csv("C:/Users/WeiDengt/Desktop/PowerBI-Report/weekly-data/WeeklySLA_NoMet.csv",index=False)
    
    tk_NoMet.to_csv(filepath,index=False)
    
    sla=SLAPercentage(tlist, tk_NoMet)
    filepath="./weekly_workedhours/"+filedate+"_WeeklySLA_NoMet_log.csv"
    sla.to_csv(filepath,index=False)
    
    return tlist

            
    

def getAccountList():
    accountlist=pd.read_csv(CONTACKFILE, encoding='cp1252')
    
    accounts=list(set(accountlist['CUSTOMER']))
    return accounts



def getEstDateTime(timestr):
    import datetime
    import pytz
    
    e_t=datetime.datetime.strptime(timestr, '%Y-%m-%d %I:%M%p')
    #e_t=datetime.datetime.strptime(timestr, '%Y-%m-%d %I:%M%p')
    e_time=e_t.astimezone(pytz.timezone('US/Eastern'))
    
    return e_time

    

def getCreateDateTime():
    today=datetime.datetime.today()
    
    #below line only for test, using yesterday
    #today=datetime.datetime.today()-datetime.timedelta(days=1)
    timeentrystart=datetime.datetime.today()-datetime.timedelta(days=7)
#    if (today.weekday()==0):
#        lastday=today-datetime.timedelta(days=3)
#    else:
#        lastday=today-datetime.timedelta(days=1)
    
    end_str=str(today.year)+"-"+str(today.month)+"-"+str(today.day)+" 10:00pm"
    #begin_str=str(today.year)+"-"+str(today.month)+"-"+str(today.day)+" 7:00am"
    testart_str=str(timeentrystart.year)+"-"+str(timeentrystart.month)+"-"+str(timeentrystart.day)+" 02:00am"
    filedate=str(today.year)+"-"+str(today.month)+"-"+str(today.day)
    
    return testart_str, end_str, filedate

def getAccountNameByID(accountID, AccountList):
    AccountName="Empty Account Name"
    for a in AccountList:
        if str(a['id'])==str(accountID):
            AccountName=a['AccountName']
    return AccountName

def getContactName(tkid, contactlist):
    for a in contactlist:
        if int(a['id'])==int(tkid):
            try:
                contactName=a['FirstName']+' '+a['LastName']
            except:
                contactName=a['FirstName']
    return contactName

def getConnect ():
    
    at=atws.connect(username='',password='')
    return at

"""
getTicketList by date range
date format: "yyyy-mm-dd"
"""
def getTicketListByDate(conn,startDate, endDate):
    tk=atws.Query('Ticket')
    tk.WHERE('ResolvedDateTime',tk.GreaterThanorEquals,startDate)
    tk.open_bracket('AND')
    tk.WHERE('ResolvedDateTime',tk.LessThanOrEquals,endDate)
    tk.close_bracket()
    tklist=conn.query(tk).fetch_all()
    
    return tklist


"""
get picklist for names of:
    IssueType, SubIssueType,Priority,Status,QueueName by ID numbers
only for entity: Ticket
"""
def getPicklist(conn,entityName):
    plist=conn.picklist.lookup(entityName)
    return plist

"""
read following names from autotask:
    resource,account, contact
    
"""
def getAutotaskName(conn, entityName):
    r=atws.Query(entityName)
    r.WHERE('id',r.NotEqual,'')
    rlist=conn.query(r).fetch_all()
    return rlist


def SLAPercentage(tklist, tk_NoMet):
    col=[ 'AccountName', 'TotalTickets','Not_Met','SLA_Not_Met']
    temp=pd.DataFrame(columns=col)
    
    i=0
    for acc in set(tk_NoMet['AccountName']):
        name=acc
        total=len(tklist[(tklist['AccountName']==name)])
        nomet=len(tk_NoMet[(tk_NoMet['AccountName']==name)])
        percentage=nomet/total
        
        temp.loc[str(i)]=[str(name),total,nomet,percentage]
        i=i+1
    
    return temp
    
         
               
def getTicketNumberByID(conn,id):
    
    try:
        
        tk=atws.Query('Ticket')
        tk.WHERE('id',tk.Equals,int(id))
        tknumlist=conn.query(tk).fetch_one()
        tknum=tknumlist['TicketNumber']
    except:
        tknum="Manual Input"
    
    return tknum

#

if __name__=='__main__':
#    
    getSLATicketList()  


