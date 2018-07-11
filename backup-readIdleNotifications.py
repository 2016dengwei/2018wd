# -*- coding: utf-8 -*-
"""
Created on Wed Jul 11 08:52:38 2018

TemplateName =
RecipientEmailAddress=
EntityNumber=ticketnumber
EntityTitle=ticketTitle

backup version without name and pwd
"""

import pandas as pd
import atws
import datetime


class Notifications:
    def __init__(self,startdate,enddate):
        self.name=""
        self.pwd=""
        
        self.start=str(startdate)+" 12:00am"
        self.end=str(enddate)+" 6:00pm"
        
        
        
    def connect(self):
        at=atws.connect(username=self.name,password=self.pwd)
        return at
    
    def readNotification(self,conn):
        hr_delta=datetime.timedelta(hours=-9)
        begin=datetime.datetime.strptime(self.start, '%Y-%m-%d %I:%M%p')+hr_delta
        close=datetime.datetime.strptime(self.end, '%Y-%m-%d %I:%M%p')+hr_delta
        
        tk=atws.Query('NotificationHistory')
        tk.WHERE('NotificationSentTime',tk.GreaterThanorEquals,str(begin))
        tk.open_bracket('AND')
        tk.WHERE('NotificationSentTime',tk.LessThanOrEquals,str(close))
        tk.close_bracket()
        
        notifications=conn.query(tk).fetch_all()
        return notifications
    
    def asDF(self, notifications):
        col=['sentDate', 'Resource','TKNum','TKTitle','IdleHours']
        df=pd.DataFrame(columns=col)
        
        i=0
        for n in notifications:
            if ('EntityNumber' in n):
                
                hr=0

                if n.TemplateName=="NotificationByWD":
                    hr=24
                if n.TemplateName=="idleNotification2BD_WD":
                    hr=48
                if n.TemplateName=="idleNotification3BD_WD":
                    hr=72
                
                try:
                    sentDate=str(n['NotificationSentTime'].year)+str("-")+str(n['NotificationSentTime'].month)+str("-")+str(n['NotificationSentTime'].day)
                    df.loc[str(i)]=[str(sentDate),str(n['InitiatingResourceID']),str(n['EntityNumber']),str(n['EntityTitle']),hr]
                except:
                    df.loc[str(i)]=[str(sentDate),'TBC',str(n['EntityNumber']),str(n['EntityTitle']),hr]

                i=i+1
        
        df_temp=df[df['IdleHours']!=0].drop_duplicates()
        df_temp.to_csv("idle_py.csv",index=False)
        return df_temp
    
    
                
def test():
    #change date to monday of last week and sunday of last week. (new week start from monday)
    
    today=datetime.datetime.today()
    startDay=today-datetime.timedelta(days=4)

    startdate=str(startDay.year)+"-"+str(startDay.month)+"-"+str(startDay.day)
    enddate=str(today.year)+"-"+str(today.month)+"-"+str(today.day)
    #print(startDay,today, startdate,enddate)
    
    #modify documentcontrol content, update date to reporting week
    dc=pd.read_csv("c:/Users/WeiDengt/Desktop/PowerBI-Report/weekly-data/DocControlTable_weeklyTK.csv")
    dc.Value[1]=startdate+" to "+enddate
    dc.to_csv("c:/Users/WeiDengt/Desktop/PowerBI-Report/weekly-data/DocControlTable_weeklyTK.csv", index=False)
    
    noti=Notifications(startdate, enddate)
    at=noti.connect()
    nfs=noti.readNotification(at)
    nfs_df=noti.asDF(nfs)
    
    return nfs_df             

if __name__=='__main__':
#    
    test() 