# -*- coding: utf-8 -*-
"""
Created on Sun Mar 02 00:00:41 2014

@author: 管理员
"""
#%%
import win32com.client
from win32com.gen_py import msof,mspp

#%% 新建一个测试用的ppt
application = win32com.client.Dispatch("PowerPoint.Application")
application.Visible = True

#%%
filename = r'c:\sample1.ppt'
presentation = application.Presentations.Open(filename)

#%%
slide = presentation.Slides[0]
#%%
shapes = slide.Shapes
len(shapes)
    
#%%    
comments = slide.Comments
len(comments)

#%%