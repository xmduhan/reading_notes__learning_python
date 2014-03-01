# -*- coding: utf-8 -*-
"""
Created on Sat Mar 01 17:48:16 2014

@author: 管理员
"""

#%% 导入相关库
import win32com.client

#%% 新建一个测试用的ppt
application = win32com.client.Dispatch("PowerPoint.Application")
application.Visible = True
presentation = application.Presentations.Add()
presentation.Slides.Add(1,1)
presentation.Slides.Add(2,2)
presentation.Slides.Add(3,3)


