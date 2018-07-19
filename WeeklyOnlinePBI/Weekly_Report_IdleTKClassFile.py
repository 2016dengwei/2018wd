# -*- coding: utf-8 -*-
"""
part of weekly report file
will import to weekly report online file

@author: WeiDengt
"""

import atws
import pandas as pd
import datetime


class Notifications:
    def __init__(self,startdate,enddate,toLondon):
        
        self.start=startdate
        self.end=enddate
        self.timeDiff=toLondon

    
    def readNotification(self,conn):
#        hr_delta=datetime.timedelta(hours=-9)
#        begin=datetime.datetime.strptime(self.start, '%Y-%m-%d %I:%M%p')+hr_delta
#        close=datetime.datetime.strptime(self.end, '%Y-%m-%d %I:%M%p')+hr_delta
#        
        tk=atws.Query('NotificationHistory')
        tk.WHERE('NotificationSentTime',tk.GreaterThanorEquals,self.start)
        tk.open_bracket('AND')
        tk.WHERE('NotificationSentTime',tk.LessThanOrEquals,self.end)
        tk.close_bracket()
        
        notifications=conn.query(tk).fetch_all()
        return notifications
    
    def asDF(self,conn):
        notifications=self.readNotification(conn)
        
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
                    notiTime=n['NotificationSentTime']+datetime.timedelta(hours=self.timeDiff)
                    sentDate=str(notiTime.year)+str("-")+str(notiTime.month)+str("-")+str(notiTime.day)
                    df.loc[str(i)]=[str(sentDate),str(n['InitiatingResourceID']),str(n['EntityNumber']),str(n['EntityTitle']),hr]
                except:
                    df.loc[str(i)]=[str(sentDate),'TBC',str(n['EntityNumber']),str(n['EntityTitle']),hr]

                i=i+1
        
        df_temp=df[df['IdleHours']!=0].drop_duplicates()
        #df_temp.to_csv("idle_py.csv",index=False)
        return df_temp