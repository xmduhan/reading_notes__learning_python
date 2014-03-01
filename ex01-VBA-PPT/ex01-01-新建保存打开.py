# -*- coding: utf-8 -*-
"""
Created on Sat Mar 01 15:15:59 2014

@author: 管理员
"""

#%% 导入相关库文件
import win32com.client, sys

#%% 启动PowerPoint
Application = win32com.client.Dispatch("PowerPoint.Application")
Application.Visible = True

#%% 新增一个PPT文件,相当于ppt中的按了"新建"菜单
Presentation = Application.Presentations.Add()
# Slides.Add(PageIndex=页面插入位置,LayoutIndex=版式编号)
Slide = Presentation.Slides.Add(1,1)
Slide = Presentation.Slides.Add(2,2)
Slide = Presentation.Slides.Add(3,3)

#%% 保存到文件
filename = r'c:\1.ppt'
Presentation.SaveAs(filename)
Presentation.Close()

#%% 从一个路径打开文件
Presentation = Application.Presentations.Open(filename)



