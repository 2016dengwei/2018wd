# -*- coding: utf-8 -*-
"""
use by weekly report, upload  data to sql express

@author: WeiDengt
"""
import pyodbc
import datetime

class ToMSSQL:
    def __init__(self):
        self.conn=pyodbc.connect('Driver={SQL Server};'
                         'Server=AES-RPT01\SQLEXPRESS;'
                         'Database=WeeklyTKSummary;'
                         'Trusted_Connection=yes;')
    
    def insertCompletedTK(self,df):
                
        cursor=self.conn.cursor()
        
        #del old records before insert new records
        cursor.execute("delete from dbo.CompletedTickets")
        
        for i in range(len(df)):
            #replace "'" and """ to "" before insert into database table
            
            title=str(df['TicketTitle'][i]).replace("'","")
            title=title.replace('"','')
            values_str="('"+df['FullName'][i]+"','"+df['TicketNumber'][i]+"','"+title+"','"+str(df['IssueType'][i])+"','"+df['SubIssueType'][i]+"','"+str(df['CreateDateTime'][i])+"','"+str(df['FirstAssignedDateTime'][i])+"','"+str(df['SLAFirstResponseDateTime'][i])+"','"+str(df['SLAResolvedDateTime'][i])+"','"+str(df['Status'][i])+"','"+str(df['Account'][i])+"','"+str(df['Queue'][i])+"')"
            sqlstr="insert into [dbo].[CompletedTickets] ([FullName],[TicketNumber],[TicketTitle],[IssueType],[SubIssueType],[CreateDateTime],[FirstAssignedDateTime],[SLAFirstResponseDateTime],[SLAResolvedDateTime],[Status],[Account],[Queue]) VALUES " +values_str 
            cursor.execute(sqlstr)
        
        
        self.conn.commit()
        
        #self.conn.close()

    
    def insertAssignedTK(self,df):

        cursor=self.conn.cursor()
        
        #del old records before insert new records
        cursor.execute("delete from AssignedTickets")

        for i in range(len(df)):
            
            title=str(df['TicketTitle'][i]).replace("'","")
            title=title.replace('"','')
            
            values_str="('"+df['FullName'][i]+"','"+df['TicketNumber'][i]+"','"+title+"','"+df['Account'][i]+"','"+str(df['SLAStartDateTime'][i])+"','"+str(df['SLAFirstResponseDateTime'][i])+"','"+str(df['FirstAssignedDateTime'][i])+"','"+df['IssueType'][i]+"','"+df['SubIssueType'][i]+"','"+str(df['Status'][i])+"','"+str(df['Queue'][i])+"')"
            sqlstr="insert into [dbo].[AssignedTickets] ([FullName],[TicketNumber],[TKTitle],[AccountName],[SLAStartDateTime],[SLAFirstResponseDate],[FirstAssigned],[Issue],[SubIssue],[Status],[Queue]) VALUES" +values_str 
            cursor.execute(sqlstr)
            #print(i)
        
        self.conn.commit()
            
        #self.conn.close()
        
        
    def insertIdleTK(self,df):
        cursor=self.conn.cursor()
        
        #del old records before insert new records
        cursor.execute("delete from idleNotification")

        for i in range(len(df)):
            
            title=str(df['TKTitle'][i]).replace("'","")
            title=title.replace('"','')
            
            values_str="('"+str(df['sentDate'][i])+"','"+df['Resource'][i]+"','"+df['TKNum'][i]+"','"+title+"','"+str(int(df['IdleHours'][i]))+"')"
            sqlstr="insert into [dbo].[idleNotification] ([sentDate],[Resource],[TKNum],[TKTitle],[IdleHours]) VALUES" +values_str 
            cursor.execute(sqlstr)
            #print(i)
        
        self.conn.commit()
            
        #self.conn.close()    


    def insertSLANotMetTK(self,df):

        cursor=self.conn.cursor()
        
        #del old records before insert new records
        cursor.execute("delete from SLA_No")

        for i in range(len(df)):
            
            values_str="('"+df['TKNumber'][i]+"','"+df['AccountName'][i]+"','"+df['Priority'][i]+"','"+df['Status'][i]+"','"+df['ContactName'][i]+"','"+df['Source'][i]+"','"+df['Issue'][i]+"','"+df['SubIssue'][i]+"','"+str(df['FRDateTime'][i])+"','"+str(df['FRDueDateTime'][i])+"','"+str(int(df['FR_SLAMet'][i]))+"','"+str(df['SLVDateTime'][i])+"','"+str(df['SLVDueDateTime'][i])+"','"+str(int(df['Actual SLA Met Tickets'][i]))+"','"+str(df['QueueID'][i])+"','"+str(int(df['Final'][i]))+"')"
            sqlstr="insert into [dbo].[SLA_No] ([TKNumber],[AccountName],[Priority],[Status],[ContactName],[Source],[Issue],[SubIssue],[FRDateTime],[FRDueDateTime],[FR_SLAMet],[SLVDateTime],[SLVDueDateTime],[ActualSLAMetTickets],[QueueID],[Final]) VALUES" +values_str 
            cursor.execute(sqlstr)
            #print(i)
        
        self.conn.commit()
            
        #self.conn.close()


    def insertWorkHoursTK(self,df):
        
        cursor=self.conn.cursor()
        
        #del old records before insert new records
        cursor.execute("delete from WorkedHours")

        for i in range(len(df)):
            
            title=str(df['TKTitle'][i]).replace("'","")
            title=title.replace('"','')
            
            values_str="('"+df['FullName'][i]+"','"+str(df['HoursWorked'][i])+"','"+df['ProjectName'][i]+"','"+df['ProjectStatus'][i]+"','"+str(df['AccountName'][i])+"','"+title+"','"+str(df['TKNum'][i])+"','"+str(df['WorkedDate'][i])+"','"+df['Queue'][i]+"','"+df['Issue'][i]+"','"+str(df['SubIssue'][i])+"')"
            sqlstr="insert into [dbo].[WorkedHours] ([FullName],[HoursWorked],[ProjectName],[ProjectStatus],[AccountName],[TaskorTicketTitle],[TaskTicketNumber],[WorkedDate],[QueueName],[Issue],[SubIssue]) VALUES" +values_str 
            cursor.execute(sqlstr)
            #print(i)
        
        self.conn.commit()
    
    def updateDocControl(self,lastUpdate,timeDiff):
        
        cursor=self.conn.cursor()
        
        #del old records before insert new records
        cursor.execute("delete from DocControl")
        
        values_str="('Title: ','Ticket Summary for Current week')"
        sqlstr="insert into [dbo].[DocControl] ([DocControlItem],[Value]) VALUES" +values_str                
        cursor.execute(sqlstr)
        
        updatedAt=lastUpdate+datetime.timedelta(hours=timeDiff)
        last=str(updatedAt.year)+"-"+str(updatedAt.month)+"-"+str(updatedAt.day)+" "+str(updatedAt.hour)+":00"
        values_str="('Last Update:','"+str(last)+"')"
        sqlstr="insert into [dbo].[DocControl] ([DocControlItem],[Value]) VALUES" +values_str                
        cursor.execute(sqlstr)
        self.conn.commit()
        #self.conn.close()

##    
#    def closeConn(self):
#        self.conn.close()
        
