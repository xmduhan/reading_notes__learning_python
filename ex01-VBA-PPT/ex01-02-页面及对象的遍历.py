# -*- coding: utf-8 -*-
"""
Created on Sat Mar 01 17:24:36 2014

@author: 管理员
"""

#%% 导入相关库
import win32com.client

#%% 新建一个测试用的ppt
application = win32com.client.Dispatch("PowerPoint.Application")
application.Visible = True
presentation = Application.Presentations.Add()
presentation.Slides.Add(1,1)
presentation.Slides.Add(2,2)
presentation.Slides.Add(3,3)

#%% 检查幻灯片的页数
slides = presentation.Slides
len(slides)

#%% 遍历幻灯片的每1页
for page in slides : print(page)

#%% 提取每1业作为变量
page1 = slides[0]
page2 = slides[1]
page3 = slides[2]

#%% 查看页面中的对象个数
shapes1 = page1.Shapes
shapes2 = page2.Shapes
shapes3 = page3.Shapes
len(shapes1)

#%% 遍历每一个对象
for shape in shapes1:
    print(shape)

#%%
for shape in shapes2:
    print(shape)

#%%
for shape in shapes3:
    print(shape)

