# -*- coding: utf-8 -*-
"""
Created on Wed Jul 18 10:16:53 2018

will import to weekly report online file
class for create worked hours df

@author: WeiDengt
"""
import pandas as pd
import atws
import datetime


class WorkedHours:
    def __init__(self,startDate, endDate,toLondon):
        
        self.start=startDate
        self.end=endDate
        self.timeDiff=toLondon
        
    def getWorkHours(self,conn):
        #conn=self.getConnect()
        tk=atws.Query('TimeEntry')
        tk.WHERE('DateWorked',tk.GreaterThanorEquals,self.start)
        tk.open_bracket('AND')
        tk.WHERE('DateWorked',tk.LessThanOrEquals,self.end)
        tk.close_bracket()
        tklist=conn.query(tk).fetch_all()
        
        return tklist
    
    def getTKInfo(self,conn,tkid):
        
        tk=atws.Query('Ticket')
        tk.WHERE('id',tk.Equals,tkid)

        ticket=conn.query(tk).fetch_one()
        
        return ticket
    
    def getPicklist(self,conn,entityName):
        plist=conn.picklist.lookup(entityName)
        return plist
    
    """
    read following names from autotask:
        resource,account, contact
        
    """
    def getAutotaskName(self,conn, entityName):
        r=atws.Query(entityName)
        r.WHERE('id',r.NotEqual,'')
        rlist=conn.query(r).fetch_all()
        return rlist
    
    def getAccountNameByID(self,accountID, AccountList):
        AccountName="Empty Account Name"
        for a in AccountList:
            if str(a['id'])==str(accountID):
                AccountName=a['AccountName']
        return AccountName
    
    def getResourceNameByID(self,resourceID, ResourceList):
        Name="NotAssigned"
        for a in ResourceList:
            if str(a['id'])==str(resourceID):
                Name=a['UserName']
        return Name
    
    
    def getWorkedHoursDF(self,conn):

        whlist=self.getWorkHours(conn)
        
        col=['FullName','HoursWorked','ProjectName','ProjectStatus','AccountName','TKTitle','TKNum','WorkedDate','Queue','Issue','SubIssue']
        whdf=pd.DataFrame(columns=col)
        
        #get picklist for name of issue, subissue, status, priority, ticketsource
        plist=self.getPicklist(conn,'Ticket')
        
        #get account and resource list from autotask
        accountlist=self.getAutotaskName(conn,'Account')
        resourcelist=self.getAutotaskName(conn,'Resource')
        
        i=0
        for wh in whlist:
            try:
                ticket=self.getTKInfo(conn,wh['TicketID'])
                tknum=ticket['TicketNumber']
                
                try:
                    tktitle=ticket['Title']
                except:
                    tktitle=""
                
                try:
                    account=self.getAccountNameByID(str(ticket['AccountID']),accountlist)
                except:
                    account=""
                    
                
                try:
                    fullname=self.getResourceNameByID(str(wh['ResourceID']),resourcelist)
                except:
                    fullname=""
                
                try:
                    hrs=wh['HoursWorked']
                except:
                    hrs=0
                                
                try:
                    wdate=str(wh['EndDateTime']+datetime.timedelta(hours=self.timeDiff))
                    
                except:
                    wdate=""
                
                
                try:
                    queue=plist['QueueID'].reverse_lookup(str(ticket['QueueID']))
                except:
                    queue=""
                
                try:
                    issue=plist['IssueType'].reverse_lookup(str(ticket['IssueType']))
                except:
                    issue=""
                
                try:
                    subissue=plist['SubIssueType'].reverse_lookup(str(ticket['SubIssueType']))
                except:
                    subissue=""
                
                
                projName=""
                projstatus=""
               
                whdf.loc[str(i)]=[fullname,hrs,projName,projstatus,account,tktitle,tknum,str(wdate),queue,issue,subissue]
                
                i=i+1
            except:
                print("")
        
        
        return whdf