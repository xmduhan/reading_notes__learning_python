# -*- coding: utf-8 -*-
"""
Created on Mon Mar 10 11:05:58 2014

@author: 14F
"""

#%% 导入必要的库
from pandas import DataFrame
from pandas.io.excel import ExcelWriter
import win32com.client
from win32com.gen_py import msof,mspp,msxl
from string import uppercase
from pandas import Series

#%% 将常量发布到全局命名空间中去
g = globals()
for c in dir(msof.constants) : g[c] = getattr(msof.constants, c)
for c in dir(mspp.constants) : g[c] = getattr(mspp.constants, c)
for c in dir(msxl.constants) : g[c] = getattr(msxl.constants, c)    

#%% 生成行名和坐标的对应关系表
luc  = list(uppercase)
columns = Series((luc + [i+j for i in luc for j in luc])[:256],range(1,257))
def cellName(nRow,nCol):
    return columns[nCol]+str(nRow)


#%%
def render(sheet,chart=None):
    # 获取数据范围
    nRow = sheet.UsedRange.Rows.Count  
    nCol = sheet.UsedRange.Columns.Count
    # 为嵌入式图表计算
    chartObjectXCells = 10
    chartObjectYCells = 25
    chartObjectLeft = sheet.Cells(2,nCol+2).Left
    chartObjectTop = sheet.Cells(2,1).Top
    chartObjectWidth = sheet.Cells(2,nCol+2+chartObjectXCells).Left - chartObjectLeft
    chartObjectHeight = sheet.Cells(2+chartObjectYCells,1).Top - chartObjectTop

    # 如果没有提供图表则要新建
    if not chart :
        chartObject = sheet.ChartObjects().Add(
        chartObjectLeft,chartObjectTop,chartObjectWidth,chartObjectHeight)
        chart = chartObject.Chart
        
    # 设置图表类型
    chart.ChartType = xlColumnClustered    

    # 增加系列
    seriesCollection = chart.SeriesCollection()
    for i in range(2,nCol+1):
        rangeName = cellName(2,i) + ':' + cellName(nRow,i)
        series = seriesCollection.NewSeries()    
        series.Name = sheet.Cells(1,i)
        series.Values = sheet.Range(rangeName) 
    # 设置x轴
    xRangeName = cellName(2,1) + ':' + cellName(nRow,1)     
    seriesCollection.Item(1).XValues = sheet.Range(xRangeName)    
        
    # 设置背景颜色
    chart.ChartArea.Interior.ColorIndex = 0
    chart.PlotArea.Interior.ColorIndex = 0
    
    # 设置背景颜
    chart.ChartArea.Interior.Color = 0xffffff
    
    # 设置图表边框的颜色
    chart.ChartArea.Border.ColorIndex = xlColorIndexNone
    
    # 设置绘图区边框
    chart.PlotArea.Border.ColorIndex = xlColorIndexNone
    
    # 设置图例位置
    chart.Legend.Position = xlLegendPositionTop
    
    # 设置x和y坐标上的主要网格线
    chart.Axes().Item(1).HasMajorGridlines = False
    chart.Axes().Item(2).HasMajorGridlines = False
    # 设置x和y坐标上的次要网格线
    chart.Axes().Item(1).HasMinorGridlines = False
    chart.Axes().Item(2).HasMinorGridlines = False
    
    # 调整绘图区的位置
    chartPlotAreaHeight = chart.PlotArea.Height + chart.PlotArea.Top
    chart.PlotArea.Top = 0
    chart.PlotArea.Height = chartPlotAreaHeight
    #
    #chartPlotAreaWidth = chart.PlotArea.Width + chart.PlotArea.Left
    #chart.PlotArea.Left = 0
    #chart.PlotArea.Width = chartPlotAreaWidth
    
    # 修改坐标轴的字体 
    for axes in chart.Axes() :
        axes.TickLabels.Font.Size=8

    # 去掉标题
    chart.HasTitle = False
    
    return chart
