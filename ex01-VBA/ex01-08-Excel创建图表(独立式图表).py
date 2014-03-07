# -*- coding: utf-8 -*-
"""
Created on Thu Mar 06 19:17:49 2014

@author: 管理员
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
    
#%% 生成一份测试excel数据文件
filename = r'c:\test1.xls'
sheetname = 'sheet1'
data = DataFrame(
    {'a':range(10), 'b':range(10,20), 'c':range(20,30)},
    index=list(uppercase)[:10]
)
datafile = ExcelWriter(filename)
data.to_excel(datafile,sheetname)
datafile.save()

#%% 使用VBA将其数据文件打开
application = win32com.client.Dispatch('Excel.Application')
application.Visible = True
application.DisplayAlerts = False
workbook = application.Workbooks.Open(filename)
sheets = workbook.Sheets 
sheet = sheets.Item(sheetname)

#%% 获取数据范围
nRow = sheet.UsedRange.Rows.Count  
nCol = sheet.UsedRange.Columns.Count
# 为嵌入式图表计算
chartObjectXCells = 10
chartObjectYCells = 25
chartObjectLeft = sheet.Cells(2,nCol+2).Left
chartObjectTop = sheet.Cells(2,1).Top
chartObjectWidth = sheet.Cells(2,nCol+2+chartObjectXCells).Left - chartObjectLeft
chartObjectHeight = sheet.Cells(2+chartObjectYCells,1).Top - chartObjectTop


#%% 使用独立式图表还是嵌入式图表 
separatedChart = False
if separatedChart :
    chart = workbook.Charts.Add()
else :
    chartObject = sheet.ChartObjects().Add(
        chartObjectLeft,chartObjectTop,chartObjectWidth,chartObjectHeight)
    chart = chartObject.Chart
    
#%% 设置图表类型
chart.ChartType = xlLine

#%% 增加系列
seriesCollection = chart.SeriesCollection()
for i in range(2,nCol+1):
    rangeName = cellName(2,i) + ':' + cellName(nRow,i)
    series = seriesCollection.NewSeries()    
    series.Name = sheet.Cells(1,i)
    series.Values = sheet.Range(rangeName) 
#%% 设置x轴
xRangeName = cellName(2,1) + ':' + cellName(nRow,1)     
seriesCollection.Item(1).XValues = sheet.Range(xRangeName)    
    
#%% 设置背景颜色
chart.ChartArea.Interior.ColorIndex = 0
chart.PlotArea.Interior.ColorIndex = 0

#%% 也可以使用 
chart.ChartArea.Interior.Color = 0xffffff

#%% 设置图表边框的颜色
chart.ChartArea.Border.ColorIndex = xlColorIndexNone
#chart.ChartArea.Border.ColorIndex = 5
#%% 设置绘图区边框
chart.PlotArea.Border.ColorIndex = xlColorIndexNone
#chart.PlotArea.Border.ColorIndex = 1


#%% 去掉图例
chart.HasLegend = False
#%% 回复图例
chart.HasLegend = True
#%% 设置图例位置
chart.Legend.Position = xlLegendPositionTop
''' 图例位置可选值 
xlLegendPositionCorner,
xlLegendPositionRight,
xlLegendPositionTop 
xlLegendPositionBottom,
xlLegendPositionLeft 
'''
# 也可以指定位置
#chart.Legend.Top = 0
#chart.Legend.Left = 0

#%% 是否显示坐标轴 (x,y)
chart.HasAxis = (True,True)
#%% 是否显示数据表
chart.HasDataTable = False

#%% 设置x和y坐标上的主要网格线
chart.Axes().Item(1).HasMajorGridlines = False
chart.Axes().Item(2).HasMajorGridlines = False
#%% 设置x和y坐标上的次要网格线
chart.Axes().Item(1).HasMinorGridlines = False
chart.Axes().Item(2).HasMinorGridlines = False

#%% 调整绘图区的位置
chartPlotAreaHeight = chart.PlotArea.Top+chart.PlotArea.Height
chart.PlotArea.Top = 0
chart.PlotArea.Height += chartPlotAreaHeight

#%% 另存文件
filename = r'c:\test2.xls'
workbook.SaveAs(filename)  
workbook.Close()  
application.Quit()