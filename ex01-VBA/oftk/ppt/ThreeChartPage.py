# -*- coding: utf-8 -*-
"""
Created on Sun Mar 02 21:55:41 2014

@author: 管理员
"""

#%% 导出模块和VB常量
from __future__ import division
import win32com.client
from win32com.gen_py import msof,mspp
#from xlchart import Line,ColumnStacked,ColumnClustered
from oftkl import clearSheet,copySheet,df2sheet

#%%
g = globals()
for c in dir(msof.constants) : g[c] = getattr(msof.constants, c)
for c in dir(mspp.constants) : g[c] = getattr(mspp.constants, c)    
    
#%%

def render(slide,mainTitle,titles,dataFrames,drawers):
    
    # 删除所有页面上的对象
    for i in [s for s in slide.Shapes] : i.Delete()
    # 注意这里不能写成 : for s in slide1.Shapes : s.Delete()
    # 因为s.Delete() 后 slide1.Shapes 就少了一个元素，和MT4中出现的问题相同
    # 先转化为Python的list后解决这个问题
    shapes = slide.Shapes        
   
    # 计算所有对象的位置参数
    border = 8
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
    #
    bigTitleBarWidth = titleBarWidth * 2 + border
    bigChartWidth = chartWidth * 2 + border
       
    # 添加顶部标题栏
    titleFrame = shapes.AddShape(msoShapeRectangle,0,0,pageWidth,titleFrameHeight)
    # 字体和段段落属性
    titleFrame.Line.Transparency = 1                      # 设置为无边框
    titleFrame.Fill.Transparency = 1                      # 设置为无填充色
    titleTextRange = titleFrame.TextFrame.TextRange
    titleTextRange.Text = mainTitle.encode('gbk')
    titleFont = titleTextRange.Font
    titleFont.Bold = msoTrue
    titleFont.Size = 28
    titleFont.Color.RGB = 0                              # 文字设置为黑色     
    titleTextRange.ParagraphFormat.Alignment = ppAlignLeft
   
    # 添加矩形标题栏
    titleBar1 = shapes.AddShape(msoShapeRectangle, titleBarBeginV1, titleBarBeginH1, titleBarWidth, titleBarHeight)
    titleBar2 = shapes.AddShape(msoShapeRectangle, titleBarBeginV2, titleBarBeginH1, titleBarWidth, titleBarHeight)
    titleBar3 = shapes.AddShape(msoShapeRectangle, titleBarBeginV1, titleBarBeginH2, bigTitleBarWidth, titleBarHeight)
    titleBars = [titleBar1,titleBar2,titleBar3]    
    
    for (i,titleBar) in enumerate(titleBars):
        titleBar.TextFrame.TextRange.Text = titles[i].encode('gbk')
        titleBar.Line.Transparency = 1                    # 设置为无边框
        titleBar.Fill.Transparency = 1                    # 设置为无填充色
        TextRange = titleBar.TextFrame.TextRange
        TextRange.Font.Color.RGB = 0                      # 字体设置为黑色
        TextRange.Font.Size = 14                          # 字体大小设置为14
        TextRange.Font.Bold = msoTrue                     # 设置为粗体字
        TextRange.ParagraphFormat.Alignment = ppAlignLeft # 设置左对齐
   
    # 添加excel图表
    shape1 = shapes.AddOLEObject(chartBeginV1, chartBeginH1, chartWidth, chartHeight,'Excel.Chart')
    shape2 = shapes.AddOLEObject(chartBeginV2, chartBeginH1, chartWidth, chartHeight,'Excel.Chart')
    shape3 = shapes.AddOLEObject(chartBeginV1, chartBeginH2, bigChartWidth, chartHeight,'Excel.Chart')
    shapeList = [shape1,shape2,shape3]
    # 找出shape中的chart和sheet对象
    charts=[]
    pptSheets=[]
    for shape in shapeList :
        workBook = shape.OLEFormat.Object 
        charts.append(workBook.Charts.Item(1))
        pptSheets.append(workBook.Sheets.Item('sheet1'))

    # 删除图标中的系列
    for chart in charts :
        for i in [s for s in chart.SeriesCollection()] : i.Delete()
            
    # 把数据拷贝入ppt并重新画图
    for (i,pptSheet) in enumerate(pptSheets) :
        #clearSheet(pptSheet)
        #copySheet(sheets[i],pptSheet)
        df2sheet(dataFrames[i],pptSheet)
        drawers[i].render(pptSheets[i],charts[i])
        shapeList[i].OLEFormat.Object 

    #%% touch一下OLE Object强迫其进行刷新 
    for shape in shapeList: shape.OLEFormat.Object 

