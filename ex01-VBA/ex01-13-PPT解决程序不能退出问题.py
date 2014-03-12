# -*- coding: utf-8 -*-
"""
Created on Wed Mar 12 18:54:15 2014

@author: 管理员
"""
#%% 定义关闭进程的函数
from wmi import WMI
def terminateProcess(processName):
    for i in WMI().Win32_Process(caption=processName):
        i.Terminate()

#%% 
import win32com.client, sys
application = win32com.client.Dispatch("PowerPoint.Application")
application.Visible = True
presentation = application.Presentations.Add()
presentation.SaveAs('c:\\1.ppt')
presentation.Close()
application.Quit()
terminateProcess('POWERPNT.EXE') # 解决程序不能退出的问题
        