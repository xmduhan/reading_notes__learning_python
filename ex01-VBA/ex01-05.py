# -*- coding: utf-8 -*-
"""
Created on Sun Mar 02 21:55:41 2014

@author: 管理员
"""

#%% 导出模块和VB常量
from __future__ import division
import win32com.client
from win32com.gen_py import msof,mspp
g = globals()
for c in dir(msof.constants) : g[c] = getattr(msof.constants, c)
for c in dir(mspp.constants) : g[c] = getattr(mspp.constants, c)
    
#%% 新建一份ppt文档
application = win32com.client.Dispatch("PowerPoint.Application")
application.Visible = True
presentation = application.Presentations.Add()
presentation.Slides.Add(1,1)
slides = presentation.Slides
slide1 = slides.Item(1)
shapes = slide1.Shapes
#%% 删除所有页面上的对象
for i in [s for s in slide1.Shapes] : i.Delete()
# 注意这里不能写成 : for s in slide1.Shapes : s.Delete()
# 因为s.Delete() 后 slide1.Shapes 就少了一个元素，和MT4中出现的问题相同
# 先转化为Python的list后解决这个问题
   
#%%
border = 10
pageWidth = 720   
pageHeight = 540
titleBarWidth = (pageWidth - border * 3)/2
titleBarHeight = 30
titleFrameHeight = 60
chartWidth = (pageWidth - border * 3)/2
chartHeight = (pageHeight - titleFrameHeight - titleBarHeight * 2 - border * 5)/2
titleBarBeginV1 = border
titleBarBeginV2 = border * 2 + titleBarWidth
titleBarBeginH1 = titleFrameHeight + border
titleBarBeginH2 = titleFrameHeight + titleBarHeight + chartHeight + border * 3  

chartBeginV1 = border
chartBeginV2 = border * 2 + titleBarWidth
chartBeginH1 = titleFrameHeight + titleBarHeight + border * 2
chartBeginH2 = titleFrameHeight + titleBarHeight * 2 + chartHeight + border * 4  


#%%   
shapes.AddShape(msoShapeRectangle, titleBarBeginV1, titleBarBeginH1, titleBarWidth, titleBarHeight)
shapes.AddShape(msoShapeRectangle, titleBarBeginV2, titleBarBeginH1, titleBarWidth, titleBarHeight)
shapes.AddShape(msoShapeRectangle, titleBarBeginV1, titleBarBeginH2, titleBarWidth, titleBarHeight)
shapes.AddShape(msoShapeRectangle, titleBarBeginV2, titleBarBeginH2, titleBarWidth, titleBarHeight)
shapes.AddShape(msoShapeRectangle, chartBeginV1, chartBeginH1, chartWidth, chartHeight)
shapes.AddShape(msoShapeRectangle, chartBeginV2, chartBeginH1, chartWidth, chartHeight)
shapes.AddShape(msoShapeRectangle, chartBeginV1, chartBeginH2, chartWidth, chartHeight)
shapes.AddShape(msoShapeRectangle, chartBeginV2, chartBeginH2, chartWidth, chartHeight)

#%%
shapes.AddShape(msoShapeRectangle,0,0,10,185)
shapes.AddShape(msoShapeRectangle,0,0,10,540)
