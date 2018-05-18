# -*- coding: utf-8 -*-
"""
Created on 17 May 2018
input data to sql-server for closed and open tickets for current day only
powerbi will display graphs based on these two tables.

@author: Wei Deng
"""


import atws
import pandas as pd
import datetime
import matplotlib.pyplot as plt
import pyodbc
        
class ConnectToAutotask:
       
        self.name=name
        self.pwd=pwd
        
        #14 hours difference between melbourne time and us/eastern time
        toEST=14
        self.endDate=datetime.datetime.now()-datetime.timedelta(hours=toEST)
        deltahours=datetime.datetime.now().hour+toEST
        self.startDate=datetime.datetime.now()-datetime.timedelta(hours=deltahours)
                
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
    
    def getEstDateTime(self,timestr):
        import datetime
        import pytz
    
        e_t=datetime.datetime.strptime(timestr, '%Y-%B-%d %I:%M%p')
    #e_t=datetime.datetime.strptime(timestr, '%Y-%m-%d %I:%M%p')
        e_time=e_t.astimezone(pytz.timezone('US/Eastern'))
    
        return e_time
    
    def getCompleteTicketListByDate(self,conn ):
               
        tk=atws.Query('Ticket')
        tk.WHERE('CompletedDate',tk.GreaterThanorEquals,self.startDate)
#        tk.open_bracket('AND')
#        tk.WHERE('CreateDate',tk.LessThanOrEquals,self.endDate)
#        tk.close_bracket()
        tklist=conn.query(tk).fetch_all()
        
        return tklist
    
    def getOpenTicketList(self,conn):
        tk=atws.Query('Ticket')
        tk.WHERE('Status',tk.NotEqual,5)
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
    
    def getClosedTicketDataFrame(self, conn,tklist):

    #create new empty dataframe
        cols=['TKNumber', 'AccountName','Issue','SubIssue','Resource','Status','Queue','CompletedDateTime','Year','Month','Day', 'WeekDay','WKofY','DayofY','CompletedHr']
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
                subIssue="TBC"
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
    
    #get ticket number
            tknum=tk['TicketNumber']   
            
            completedtime=tk['CompletedDate']+datetime.timedelta(hours=9)
            y=completedtime.year
            m=completedtime.month
            d=completedtime.day
            
            wkd=self.getDayOfWeek(y,m,d)
            woy=self.getWeekNo(y,m,d)
            doy=self.getDayOfYear(y,m,d)
                        
            completedtime=tk['CompletedDate']+datetime.timedelta(hours=9)
            hr=completedtime.hour
            
            temp.loc[str(i)]=[str(tknum),str(accountname),str(issue),str(subIssue),str(resource),str(status),str(Queue),str(completedtime),int(y),int(m),int(d),int(wkd),int(woy),int(doy),int(hr)]
            i=i+1
        
        #only return service desk queue , queueid=29683528
        #df_tlist=temp[(temp['QueueID']=='29683528')]
        
        return temp
    
    def getOpenTicketDataFrame(self, conn,tklist):

    #create new empty dataframe
        cols=['TKNumber', 'AccountName','Issue','SubIssue','Resource','Status','Queue','LastUpdate','Year','Month','Day', 'WeekDay','WKofY','DayofY','updateHr', 'DueDate','DueDay','DueMonth','DueYear','DueDayofY']
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
                subIssue="TBC"
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
    
    #get ticket number
            tknum=tk['TicketNumber']   
            
            #modify time difference 
            tkdue=tk['DueDateTime']+datetime.timedelta(hours=9)
            DueDay=int(tkdue.day)
            DueMonth=int(tkdue.month)
            DueYear=int(tkdue.year)
            DueDayofY=self.getDayOfYear(DueYear,DueMonth,DueDay)
            y,m,d=self.getYMD()
            wkd=self.getDayOfWeek(y,m,d)
            woy=self.getWeekNo(y,m,d)
            doy=self.getDayOfYear(y,m,d)
            hr=datetime.datetime.now().hour
    
            temp.loc[str(i)]=[str(tknum),str(accountname),str(issue),str(subIssue),str(resource),str(status),str(Queue),datetime.datetime.now(),int(y),int(m),int(d),int(wkd),int(woy),int(doy),int(hr), str(tkdue),int(DueDay),int(DueMonth),int(DueYear),int(DueDayofY)]
            i=i+1
        
        #only return service desk queue , queueid=29683528
        #df_tlist=temp[(temp['QueueID']=='29683528')]
        
        return temp

class ToMSSQL:
    def __init__(self):
        self.conn=pyodbc.connect('Driver={SQL Server};'
                         'Server=Gerry-Laptop\SQLEXPRESS;'
                         'Database=TBX;'
                         'Trusted_Connection=yes;')
    
    def insertCompletedTK(self,df):
        #table: [TKNum],[PrimaryResource],[Status],[LastUpdate],[Queue],[Issue],[SubIssue],[AccountName]
        #df: ['TKNumber', 'AccountName','Issue','SubIssue','Resource','Status','Queue','LastUpdate']
                
        cursor=self.conn.cursor()
        #del old records before insert new records
        cursor.execute("delete from dbo.sd_tk_realtime_monitor_complete")
        
        for i in range(len(df)):
            values_str="('"+df['TKNumber'][i]+"','"+df['Resource'][i]+"','"+df['Status'][i]+"','"+str(df['CompletedDateTime'][i])+"','"+df['Queue'][i]+"','"+df['Issue'][i]+"','"+df['SubIssue'][i]+"','"+df['AccountName'][i]+"','"+str(int(df['Year'][i]))+"','"+str(int(df['Month'][i]))+"','"+str(int(df['Day'][i]))+"','"+str(int(df['WeekDay'][i]))+"','"+str(int(df['WKofY'][i]))+"','"+str(int(df['DayofY'][i]))+"','"+str(int(df['CompletedHr'][i]))+"')"
            sqlstr="insert into [dbo].[sd_tk_realtime_monitor_complete] ([TKNum],[PrimaryResource],[Status],[CompletedDateTime],[Queue],[Issue],[SubIssue],[AccountName],[Year],[Month],[Day],[WeekDay],[WKofY],[DayofY],[CompletedHr]) VALUES" +values_str 
            cursor.execute(sqlstr)
        
            self.conn.commit()
    
    def insertOpenTK(self,df):
        #table: [TKNum],[PrimaryResource],[Status],[LastUpdate],[Queue],[Issue],[SubIssue],[AccountName]
        #table-2: DueDate , DueDay , DueMonth , DueYear , DueDayofY;

        #df: ['TKNumber', 'AccountName','Issue','SubIssue','Resource','Status','Queue','LastUpdate']
        #df-2:'DueDate', 'DueDay', 'DueMonth', 'DueYear', 'DueDayofY'
        cursor=self.conn.cursor()
        
        #del old records before insert new records
        cursor.execute("delete from sd_tk_realtime_monitor_open")

        for i in range(len(df)):
            values_str="('"+df['TKNumber'][i]+"','"+df['Resource'][i]+"','"+df['Status'][i]+"','"+str(df['LastUpdate'][i])+"','"+df['Queue'][i]+"','"+df['Issue'][i]+"','"+df['SubIssue'][i]+"','"+df['AccountName'][i]+"','"+str(int(df['Year'][i]))+"','"+str(int(df['Month'][i]))+"','"+str(int(df['Day'][i]))+"','"+str(int(df['WeekDay'][i]))+"','"+str(int(df['WKofY'][i]))+"','"+str(int(df['DayofY'][i]))+"','"+str(int(df['updateHr'][i]))+"','"+df['DueDate'][i] +"','"+str(int(df['DueDay'][i]))+"','"+ str(int(df['DueMonth'][i]))+"','"+str(int(df['DueYear'][i]))+"','"+str(int(df['DueDayofY'][i]))+"')"
            sqlstr="insert into [dbo].[sd_tk_realtime_monitor_open] ([TKNum],[PrimaryResource],[Status],[LastUpdate],[Queue],[Issue],[SubIssue],[AccountName],[Year],[Month],[Day],[WeekDay],[WKofY],[DayofY],[updateHr],[DueDate],[DueDay],[DueMonth],[DueYear], [DueDayofY]) VALUES" +values_str 
            cursor.execute(sqlstr)
            #print(i)
        
            self.conn.commit()
            
 
    
    def readTable(self):
        cursor=self.conn.cursor()
        sqlstr="select * from [dbo].[sd_tk_realtime_monitor]"
        #sqlstr="select * from [dbo].[testTable]"
        cursor.execute(sqlstr)
        
        rows=cursor.fetchall()
        for row in rows:
            if not row:
                break
        #self.conn.close()
    
    def closeConn(self):
        self.conn.close()
        


def sd_test():
    tk=ConnectToAutotask()
    conn=tk.getConnect()
    t_c=tk.getCompleteTicketListByDate(conn)
    t_o=tk.getOpenTicketList(conn)
    df_c=tk.getClosedTicketDataFrame(conn,t_c)
    df_o=tk.getOpenTicketDataFrame(conn,t_o)
    
    s=ToMSSQL()
    s.insertCompletedTK(df_c)
    s.insertOpenTK(df_o)
    s.closeConn()


#if __name__=='__main__':
###    
#    sd_test()  
   
    
