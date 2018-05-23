# -*- coding: utf-8 -*-
"""
Created on 23 May 2018
input data to sql-server for closed and open tickets for weekly summary reprot
powerbi will display graphs based on these 6 tables.

@author: Wei Deng
"""


import pandas as pd
import datetime
import pyodbc
import tkinter as tk
from tkinter import filedialog

        

        
class ReadFile:
    def __init__(self,y,m,d):
        #y: year, m: month, d: day
        self.y=y
        self.m=m
        self.d=d
    
   
    def getDayOfYear(self,year,month,day):
        doy=datetime.datetime(int(year),int(month),int(day)).timetuple().tm_yday
        
        return doy
    
    def getWeekNo(self,year,month,day):
        woy=datetime.datetime(int(year),int(month),int(day)).isocalendar()[1]
        
        return woy
    
    def getDayOfWeek(self,year,month,day):
        #return 0-6, present Monday (=0) to Sunday (=6)
        dow=datetime.datetime(int(year),int(month),int(day)).weekday()
        
        return dow
    
    def readSLA_F(self):
        rt=tk.Tk()
        filepath=filedialog.askopenfilename()
        filedata=pd.read_csv(filepath)  
        rt.destroy()
        
        filedata['Year']=self.y
        filedata['Month']=self.m
        filedata['Day']=self.d
        filedata['WkDay']=self.getDayOfWeek(self.y,self.m,self.d)
        filedata['WKofY']=self.getWeekNo(self.y,self.m,self.d)
        filedata['DayofY']=self.getDayOfYear(self.y,self.m,self.d) 
        
        sla_f=filedata[(filedata['Actual SLA Met Tickets']==0)]

        return sla_f
        
    def readDailyFile(self):
        rt=tk.Tk()
        filepath=filedialog.askopenfilename()
        filedata=pd.read_excel(filepath)
        filedata['Year']=self.y
        filedata['Month']=self.m
        filedata['Day']=self.d
        filedata['WkDay']=self.getDayOfWeek(self.y,self.m,self.d)
        filedata['WKofY']=self.getWeekNo(self.y,self.m,self.d)
        filedata['DayofY']=self.getDayOfYear(self.y,self.m,self.d)        
        
        rt.destroy()
        
        return filedata
    
    def readWeeklyAssignFile(self):
        rt=tk.Tk()
        filepath=filedialog.askopenfilename()
        filedata=pd.read_excel(filepath)
        rt.destroy()

        ys=[]
        ms=[]
        ds=[]
        wd=[]
        wy=[]
        dy=[]
        for i in range(len(filedata)):
            ys.append(filedata['FirstAssigned'][i].year)
            ms.append(filedata['FirstAssigned'][i].month)
            ds.append(filedata['FirstAssigned'][i].day)
            wd.append(self.getDayOfWeek(filedata['FirstAssigned'][i].year,filedata['FirstAssigned'][i].month,filedata['FirstAssigned'][i].day))
            wy.append(self.getWeekNo(filedata['FirstAssigned'][i].year,filedata['FirstAssigned'][i].month,filedata['FirstAssigned'][i].day))
            dy.append(self.getDayOfYear(filedata['FirstAssigned'][i].year,filedata['FirstAssigned'][i].month,filedata['FirstAssigned'][i].day))
        
        filedata['Year']=ys
        filedata['Month']=ms
        filedata['Day']=ds
        filedata['WkDay']=wd
        filedata['WKofY']=wy
        filedata['DayofY']=dy
        
        return filedata

    def readWeeklyCompletedFile(self):
            rt=tk.Tk()
            filepath=filedialog.askopenfilename()
            filedata=pd.read_excel(filepath)
            rt.destroy()
    
            ys=[]
            ms=[]
            ds=[]
            wd=[]
            wy=[]
            dy=[]
            for i in range(len(filedata)):
                ys.append(filedata['SLAResolvedDateTime'][i].year)
                ms.append(filedata['SLAResolvedDateTime'][i].month)
                ds.append(filedata['SLAResolvedDateTime'][i].day)
                wd.append(self.getDayOfWeek(filedata['SLAResolvedDateTime'][i].year,filedata['SLAResolvedDateTime'][i].month,filedata['SLAResolvedDateTime'][i].day))
                wy.append(self.getWeekNo(filedata['SLAResolvedDateTime'][i].year,filedata['SLAResolvedDateTime'][i].month,filedata['SLAResolvedDateTime'][i].day))
                dy.append(self.getDayOfYear(filedata['SLAResolvedDateTime'][i].year,filedata['SLAResolvedDateTime'][i].month,filedata['SLAResolvedDateTime'][i].day))
            
            filedata['Year']=ys
            filedata['Month']=ms
            filedata['Day']=ds
            filedata['WkDay']=wd
            filedata['WKofY']=wy
            filedata['DayofY']=dy
            
            return filedata

    def readWeeklyWorkedHrsFile(self):
            rt=tk.Tk()
            filepath=filedialog.askopenfilename()
            filedata=pd.read_excel(filepath)
            rt.destroy()
    
            ys=[]
            ms=[]
            ds=[]
            wd=[]
            wy=[]
            dy=[]
            for i in range(len(filedata)):
                ys.append(filedata['WorkedDate(4report)'][i].year)
                ms.append(filedata['WorkedDate(4report)'][i].month)
                ds.append(filedata['WorkedDate(4report)'][i].day)
                wd.append(self.getDayOfWeek(filedata['WorkedDate(4report)'][i].year,filedata['WorkedDate(4report)'][i].month,filedata['WorkedDate(4report)'][i].day))
                wy.append(self.getWeekNo(filedata['WorkedDate(4report)'][i].year,filedata['WorkedDate(4report)'][i].month,filedata['WorkedDate(4report)'][i].day))
                dy.append(self.getDayOfYear(filedata['WorkedDate(4report)'][i].year,filedata['WorkedDate(4report)'][i].month,filedata['WorkedDate(4report)'][i].day))
            
            filedata['Year']=ys
            filedata['Month']=ms
            filedata['Day']=ds
            filedata['WkDay']=wd
            filedata['WKofY']=wy
            filedata['DayofY']=dy
            data=filedata.fillna(value="TBC")
            
            return data




class ToMSSQL:
    def __init__(self):
        self.conn=pyodbc.connect('Driver={SQL Server};'
                         'Server=AES-RPT01\SQLEXPRESS;'
                         'Database=TBX;'
                         'Trusted_Connection=yes;')
    
    def insertDailyOpen(self,df):
        cursor=self.conn.cursor()
        
        for i in range(len(df)):
            values_str="('"+df['TicketNumber'][i]+"','"+df['FullName'][i]+"','"+df['Status'][i]+"','"+df['IssueType'][i]+"','"
            values_str=values_str+df['Sub-IssueType'][i]+"','"+str(df['Year'][i])+"','"+str(df['Month'][i])+"','"+str(df['Day'][i])+"','"
            values_str=values_str+str(df['WkDay'][i])+"','"+str(df['WKofY'][i])+"','"+str(df['DayofY'][i])+"')"
            sqlstr="insert into [dbo].[sd_tk_wk_dailyOpen] ([TKNum],[PrimaryResource],[Status],[Issue],[SubIssue],[Year],[Month],[Day],[WeekDay],[WKofY],[DayofY]) VALUES" +values_str 
            cursor.execute(sqlstr)
            
            self.conn.commit()
    
    def insertDailyClosed(self,df):
        cursor=self.conn.cursor()
        
        temp="no date"
        for i in range(len(df)):
            values_str="('"+df['TicketNumber'][i]+"','"+df['FullName'][i]+"','"+str(temp)+"','"+df['IssueType'][i]+"','"
            values_str=values_str+df['Sub-IssueType'][i]+"','"+str(df['Year'][i])+"','"+str(df['Month'][i])+"','"+str(df['Day'][i])+"','"
            values_str=values_str+str(df['WkDay'][i])+"','"+str(df['WKofY'][i])+"','"+str(df['DayofY'][i])+"')"
            sqlstr="insert into [dbo].[sd_tk_wk_SameDayCompleted] ([TKNum],[PrimaryResource],[CompletedDateTime],[Issue],[SubIssue],[Year],[Month],[Day],[WeekDay],[WKofY],[DayofY]) VALUES" +values_str 
                        
            cursor.execute(sqlstr)
            
        self.conn.commit()
    
    def insertWeekAssigned(self,df):
        cursor=self.conn.cursor()
        
        for i in range(len(df)):
            values_str="('"+df['TicketNumber'][i]+"','"+df['FullName'][i]+"','"+str(df['FirstAssigned'][i])+"','"+df['Source'][i]+"','"+df['Queue'][i]+"','"+df['Issue'][i]+"','"
            values_str=values_str+df['SubIssue'][i]+"','"+df['AccountName'][i]+"','"+str(df['Year'][i])+"','"+str(df['Month'][i])+"','"+str(df['Day'][i])+"','"
            values_str=values_str+str(df['WkDay'][i])+"','"+str(df['WKofY'][i])+"','"+str(df['DayofY'][i])+"')"
            sqlstr="insert into [dbo].[sd_tk_wk_assigned] ([TKNum],[PrimaryResource],[firstAssignedDate],[Source],[Queue],[Issue],[SubIssue],[AccountName],[Year],[Month],[Day],[WeekDay],[WKofY],[DayofY]) VALUES" +values_str 
                       
            cursor.execute(sqlstr)
            
        self.conn.commit()        

    def insertWeekCompleted(self,df):
        cursor=self.conn.cursor()
        
        for i in range(len(df)):
            values_str="('"+df['TicketNumber'][i]+"','"+df['FullName'][i]+"','"+str(df['SLAResolvedDateTime'][i])+"','"+df['IssueType'][i]+"','"
            values_str=values_str+df['SubIssueType'][i]+"','"+df['account'][i]+"','"+str(df['Year'][i])+"','"+str(df['Month'][i])+"','"+str(df['Day'][i])+"','"
            values_str=values_str+str(df['WkDay'][i])+"','"+str(df['WKofY'][i])+"','"+str(df['DayofY'][i])+"')"
            sqlstr="insert into [dbo].[sd_tk_wk_Completed] ([TKNum],[PrimaryResource],[CompletedDateTime],[Issue],[SubIssue],[AccountName],[Year],[Month],[Day],[WeekDay],[WKofY],[DayofY]) VALUES" +values_str 
                       
            cursor.execute(sqlstr)
            
        self.conn.commit()
    
    def insertSLA_F(self,df):
        cursor=self.conn.cursor()
        
        for i in range(len(df)):
            values_str="('"+df['TKNumber'][i]+"','"+df['AccountName'][i]+"','"+df['Issue'][i]+"','"+df['SubIssue'][i]+"','"+str(df['Year'][i])+"','"+str(df['Month'][i])+"','"+str(df['Day'][i])+"','"
            values_str=values_str+str(df['WkDay'][i])+"','"+str(df['WKofY'][i])+"','"+str(df['DayofY'][i])+"')"
            sqlstr="insert into [dbo].[sd_tk_SLA_F] ([TKNum],[AccountName],[Issue],[SubIssue],[Year],[Month],[Day],[WeekDay],[WKofY],[DayofY]) VALUES" +values_str 
                       
            cursor.execute(sqlstr)
            
        self.conn.commit()        
    
    
    def insertWorkedHrs(self,df):
        cursor=self.conn.cursor()
        
        for i in range(len(df)):
            values_str="('"+df['TaskTicketNumber'][i]+"','"+df['WorkEntryCreatedBy(4report)'][i]+"','"+df['Issue'][i]+"','"
            values_str=values_str+df['SubIssue'][i]+"','"+df['AccountName'][i]+"','"+str(df['WorkedDate(4report)'][i])+"','"+str(df['HoursWorked'][i])+"','"+df['QueueName'][i]+"','"+str(df['Year'][i])+"','"+str(df['Month'][i])+"','"+str(df['Day'][i])+"','"
            values_str=values_str+str(df['WkDay'][i])+"','"+str(df['WKofY'][i])+"','"+str(df['DayofY'][i])+"')"
            sqlstr="insert into [dbo].[sd_tk_workedHrs] ([TKNum],[Resource],[Issue],[SubIssue],[AccountName],[WorkedDate], [HoursWorked], [Queue],[Year],[Month],[Day],[WeekDay],[WKofY],[DayofY]) VALUES" +values_str 
                       
            cursor.execute(sqlstr)
            
        self.conn.commit()    
    
    def closeConn(self):
        self.conn.close()
        


def sd_test():
   
    
    s=ToMSSQL()
   
    s.closeConn()

 
   
    