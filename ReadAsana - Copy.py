# -*- coding: utf-8 -*-
"""
Created on Fri Jul  6 14:33:39 2018
asana personal token: "0/de6a423b34fea51491fedab6a39e5c02"
@author: Wei Deng
backup, no personal token
"""

import asana
import pandas as pd


class MyAsana:
    def __init__(self):
        self.MYTOKEN=""
        #self.WORKSPACE="Be Collective"
    
    
    def readProjects(self):
        client=asana.Client.access_token(self.MYTOKEN)
        client.options['page_size']=100
        my_client=client.users.me()
        
        #get workspaces, for BC, only 'Be Collective' exist
        workspace=my_client['workspaces']
        
        #get all projects name (such as feature 2/july-....)
        #workspace[0]['name']]='be collective'
        #each project: projet['id'], project['name']
        myprojects=client.projects.find_by_workspace(workspace[0]['id'])
        
        return client, myprojects
    
    def readTasks(self, client, projectID):
        
        #find tasks by project id
        tasks=client.tasks.find_by_project(projectID)
        
        
        return tasks
    
    
    def readTaskDetails(self, client, taskID):
        #singleTask: [task name, 'notes',project id and name, parent id and name,
        # memberships-section name is column name, ]
        #singletask['memberships'][0]['section']['name'] return column name
        
        singleTask=client.tasks.find_by_id(taskID)
        
        return singleTask
    
    
    def getDesignRef(self, tasknote):
        start=tasknote.find("http")
        end=tasknote.find("\n",start)
        
        refLink=tasknote[start:end]
        
        return refLink
    
    def getBusinessRule(self,tasknote):
        br=""
        start=tasknote.lower().find("Business Rule".lower())
        end=tasknote.lower().find("End Business Rule".lower(),start)
    
        if end>0:
            br=tasknote[start:end]
        
        return br
    
    def exportDF(self,client,tasks):
        col=['ProjectName','ColumnName','TaskName','TaskNote','DesignRef','Assignee', 'BusinessRules','CardURL']
        
        df=pd.DataFrame(columns=col)
        
        i=0
        for singletask in tasks:
            try:
                t=self.readTaskDetails(client,singletask['id'])
                pname=t['memberships'][0]['project']['name']
                colname=t['memberships'][0]['section']['name']
                taskname=t['name']
                note=t['notes']
                ref=self.getDesignRef(note)
                br=self.getBusinessRule(note)
                
                cardurl=str("https://app.asana.com/0/")+str(t['memberships'][0]['project']['id'])+"/"+str(t['id'])
                #if task not assigned yet the name is empty
                try:
                    assignee=t['assignee']['name']
                except:
                    assignee=""
                
                
                df.loc[str(i)]=[pname,colname,taskname,note,ref,assignee,br,cardurl]
                
                i=i+1
            except:
                print(i," check....")
        return df
    
    
def test():
    my=MyAsana()
    client,myp=my.readProjects()
    tasks=my.readTasks(client,"681325280041665")
    df=my.exportDF(client,tasks)
    
    return df