# -*- coding: utf-8 -*-
"""
Created on Sat Mar 01 15:15:59 2014

@author: 管理员
"""

#%% 导入相关库文件
import win32com.client, sys

#%% 启动PowerPoint
application = win32com.client.Dispatch("PowerPoint.Application")
application.Visible = True

#%% 新增一个PPT文件,相当于ppt中的按了"新建"菜单
presentation = application.Presentations.Add()
# Slides.Add(PageIndex=页面插入位置,LayoutIndex=版式编号)
presentation.Slides.Add(1,1)
presentation.Slides.Add(2,2)
presentation.Slides.Add(3,3)

#%% 保存到文件
filename = r'c:\1.ppt'
presentation.SaveAs(filename)
presentation.Close()

#%% 从一个路径打开文件
presentation = application.Presentations.Open(filename)



