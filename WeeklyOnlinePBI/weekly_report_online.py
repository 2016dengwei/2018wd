# -*- coding: utf-8 -*-
"""
Created on Mon Jul 16 13:20:56 2018

@author: WeiDengt
"""


import atws
import pandas as pd
import datetime

#import other py files
import Weekly_Report_IdleTKClassFile
import Weekly_ReportSLANotMetClassFile
import Weekly_Report_WorkedHoursClassFile
import Weekly_Report_toSQLEXPClassFile



class ConnectToAutotask:
       
    def __init__(self,name='',pwd=''):
        self.name=name
        self.pwd=pwd
        
        #14 hours difference between melbourne time and us/eastern time
        #9 hours diff from London time
        self.toLondon=9
        deltaday=datetime.datetime.now().weekday()
        self.endDate=datetime.datetime.now()-datetime.timedelta(hours=self.toLondon)
        temp_date=str(datetime.datetime.now().date()-datetime.timedelta(days=deltaday))
        temp_date_str=temp_date+" 12:00am"
        self.startDate=datetime.datetime.strptime(temp_date_str, '%Y-%m-%d %I:%M%p')-datetime.timedelta(hours=self.toLondon)
                
    def SDList(self):
        sdlist=['andrew.quach', 'julian.cini','michael.rui', 'roslyn.calvi', 'daniel.robinson','samira.javadinia', 'godfrey.yoganathan','sean.buckingham']
        
        return sdlist
    
    def getConnect(self):
        at=atws.connect(username=self.name,password=self.pwd)
        return at
    
    def getYMD(self):
        import datetime
        today=datetime.datetime.now()
        year=today.year
        month=today.month
        day=today.day
        
        return year,month,day
    
    
    def getCompleteTicketListByDate(self):
        conn=self.getConnect()
        
        tk=atws.Query('Ticket')
        tk.WHERE('CompletedDate',tk.GreaterThanorEquals,self.startDate)
        tk.open_bracket('AND')
        tk.WHERE('CreateDate',tk.LessThanOrEquals,self.endDate)
        tk.close_bracket()
        tklist=conn.query(tk).fetch_all()
        
        return tklist
    
    def getOpenTicketList(self):
        #ticket status: 5, compplete
        
        conn=self.getConnect()

        tk=atws.Query('Ticket')
        tk.WHERE('Status',tk.NotEqual,5)
        tklist=conn.query(tk).fetch_all()
        
        return tklist
    
    def getAssignedTK(self):
        conn=self.getConnect()

        tk=atws.Query('Ticket')
        tk.WHERE('FirstResponseDateTime',tk.GreaterThanorEquals,self.startDate)
        tk.open_bracket('AND')
        tk.WHERE('FirstResponseDateTime',tk.LessThanOrEquals,self.endDate)
        tk.close_bracket()
        tklist=conn.query(tk).fetch_all()
        
        return tklist


    """
    get picklist for names of:
        IssueType, SubIssueType,Priority,Status,QueueName by ID numbers
    only for entity: Ticket
    """
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
    
    
    def getDayOfYear(self,year,month,day):
        import datetime
        doy=datetime.datetime(int(year),int(month),int(day)).timetuple().tm_yday
        
        return doy
    
    def getWeekNo(self,year,month,day):
        import datetime
        woy=datetime.datetime(int(year),int(month),int(day)).isocalendar()[1]
        
        return woy
    
    def getDayOfWeek(self,year,month,day):
        #return 0-6, present Monday (=0) to Sunday (=6)
        import datetime
        dow=datetime.datetime(int(year),int(month),int(day)).weekday()
        
        return dow
        
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
    
    def getTicketDataFrame_CompleteTK(self):
        
        conn=self.getConnect()
        tklist=self.getCompleteTicketListByDate()
   #FullName,	TicketNumber,	TicketTitle,	IssueType,	SubIssueType,CreateDateTime,	FirstAssignedDateTime,	SLAFirstResponseDateTime,	SLAResolvedDateTime,	Status nvarchar(50),	Account,	Source         
        cols=['FullName','TicketNumber','TicketTitle','IssueType',	'SubIssueType','CreateDateTime','FirstAssignedDateTime','SLAFirstResponseDateTime','SLAResolvedDateTime','Status','Account','Queue' ]     
   #cols=['TKNumber', 'Title','AccountName','Issue','SubIssue','Resource','Status','Queue','LastUpdate','Year','Month','Day', 'WeekDay','WKofY','DayofY','updateHr']
        temp=pd.DataFrame(columns=cols)
 
    
    #get picklist for name of issue, subissue, status, priority, ticketsource
        plist=self.getPicklist(conn,'Ticket')
        
    #get account and resource list from autotask
        accountlist=self.getAutotaskName(conn,'Account')
        resourcelist=self.getAutotaskName(conn,'Resource')
    
        
        i=0
        for tk in tklist:
            
            issue=plist['IssueType'].reverse_lookup(str(tk['IssueType']))
            try:
                subIssue=plist['SubIssueType'].reverse_lookup(str(tk['SubIssueType']))
            except:
                subIssue='TBC'
                
            if ('Status' in tk):
                try:
                    status=plist['Status'].reverse_lookup(str(tk['Status']))
                except:
                    status='NoStatus'

                                    
            Queue='NoQueue'
            if ('QueueID' in tk):
                Queue=plist['QueueID'].reverse_lookup(str(tk['QueueID']))

    #get accountname from autotask
            accountname='empty'
            if ('AccountID' in tk):
                accountname=self.getAccountNameByID(str(tk['AccountID']),accountlist)
            
    #get resource name 
            resource="NotAssign"
            if ('AssignedResourceID' in tk):
                resource=self.getResourceNameByID(str(tk['AssignedResourceID']),resourcelist)
    
    #get ticket number and title
            title=tk['Title']
            tknum=tk['TicketNumber']   
            createDate=tk['CreateDate']+datetime.timedelta(hours=self.toLondon)
            FirstAssign=tk['FirstResponseDateTime']+datetime.timedelta(hours=self.toLondon)
            SLAFirstResponse=tk['FirstResponseDateTime']+datetime.timedelta(hours=self.toLondon)
            SLAResolved=tk['CompletedDate']+datetime.timedelta(hours=self.toLondon)
            
#            y,m,d=self.getYMD()
#            wkd=self.getDayOfWeek(y,m,d)
#            woy=self.getWeekNo(y,m,d)
#            doy=self.getDayOfYear(y,m,d)
#            hr=datetime.datetime.now().hour
    
            temp.loc[str(i)]=[str(resource),str(tknum),str(title),str(issue),str(subIssue),str(createDate),str(FirstAssign),str(SLAFirstResponse),str(SLAResolved),str(status),str(accountname),str(Queue)]
            i=i+1
        
        #only return service desk queue , queueid=29683528
        #df_tlist=temp[(temp['QueueID']=='29683528')]
        
        return temp
    
    
    def getTicketDataFrame_AssginedTK(self):
        
        conn=self.getConnect()
        
        tklist=self.getAssignedTK()
#	AccountName nvarchar(50),
#		Queue nvarchar(50),	Source nvarchar(50)

        cols=['FullName','TicketNumber','TicketTitle','Account','IssueType',	'SubIssueType','SLAStartDateTime','SLAFirstResponseDateTime','FirstAssignedDateTime','Status','Queue' ]     
        temp=pd.DataFrame(columns=cols)
 
    
    #get picklist for name of issue, subissue, status, priority, ticketsource
        plist=self.getPicklist(conn,'Ticket')
        
    #get account and resource list from autotask
        accountlist=self.getAutotaskName(conn,'Account')
        resourcelist=self.getAutotaskName(conn,'Resource')
    
        
        i=0
        for tk in tklist:
            
            issue=plist['IssueType'].reverse_lookup(str(tk['IssueType']))
            try:
                subIssue=plist['SubIssueType'].reverse_lookup(str(tk['SubIssueType']))
            except:
                subIssue='TBC'
                
            if ('Status' in tk):
                try:
                    status=plist['Status'].reverse_lookup(str(tk['Status']))
                except:
                    status='NoStatus'

                                    
            Queue='NoQueue'
            if ('QueueID' in tk):
                Queue=plist['QueueID'].reverse_lookup(str(tk['QueueID']))

    #get accountname from autotask
            accountname='empty'
            if ('AccountID' in tk):
                accountname=self.getAccountNameByID(str(tk['AccountID']),accountlist)
            
    #get resource name 
            resource="NotAssign"
            if ('AssignedResourceID' in tk):
                resource=self.getResourceNameByID(str(tk['AssignedResourceID']),resourcelist)
    
    #get ticket number and title
            title=tk['Title']
            tknum=tk['TicketNumber']   
            createDate=tk['CreateDate']+datetime.timedelta(hours=self.toLondon)
            FirstAssign=tk['FirstResponseDateTime']+datetime.timedelta(hours=self.toLondon)
            SLAFirstResponse=tk['FirstResponseDateTime']+datetime.timedelta(hours=self.toLondon)
            
#            y,m,d=self.getYMD()
#            wkd=self.getDayOfWeek(y,m,d)
#            woy=self.getWeekNo(y,m,d)
#            doy=self.getDayOfYear(y,m,d)
#            hr=datetime.datetime.now().hour
    
            temp.loc[str(i)]=[str(resource),str(tknum),str(title),str(accountname),str(issue),str(subIssue),str(createDate),str(SLAFirstResponse),str(FirstAssign),str(status),str(Queue)]
            i=i+1
        
        #only return service desk queue , queueid=29683528
        #df_tlist=temp[(temp['QueueID']=='29683528')]
        
        return temp

    def getIdleTKDF(self):
        #a=ConnectToAutotask()
        idle=Weekly_Report_IdleTKClassFile.Notifications(self.startDate,self.endDate,self.toLondon)
        conn=self.getConnect()
        idledf=idle.asDF(conn)   
        
        return idledf
    
    def getSLANotMet(self):
        notMet=Weekly_ReportSLANotMetClassFile.SLA_Record(self.startDate,self.endDate)
        conn=self.getConnect()
        notMetTKdf=notMet.getSLATicketList(conn,self.startDate,self.endDate)
        
        return notMetTKdf
    
    def getHours(self):
        whours=Weekly_Report_WorkedHoursClassFile.WorkedHours(self.startDate,self.endDate, self.toLondon)
        conn=self.getConnect()
        whdf=whours.getWorkedHoursDF(conn)   
        
        return whdf


if __name__=='__main__':

    a=ConnectToAutotask()
    df1=a.getTicketDataFrame_CompleteTK()
    df2=a.getTicketDataFrame_AssginedTK()
    df3=a.getIdleTKDF()
    df4=a.getSLANotMet()
    df5=a.getHours()
    
    
    x=Weekly_Report_toSQLEXPClassFile.ToMSSQL()
    x.insertCompletedTK(df1)
    x.insertAssignedTK(df2)
    x.insertIdleTK(df3)
    x.insertSLANotMetTK(df4)
    x.insertWorkHoursTK(df5)
    x.updateDocControl(a.endDate,a.toLondon)
    
    x.conn.close()

    


    
 

  