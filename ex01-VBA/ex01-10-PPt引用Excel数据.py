# -*- coding: utf-8 -*-
"""
Created on Fri Mar 07 23:41:05 2014

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

#%% 打开测试数据
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = True
excel.DisplayAlerts = False
workbook = excel.Workbooks.Open(filename)
sheets = workbook.Sheets 
sheet = sheets.Item(sheetname)

#%% 新建一份ppt文档
ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Visible = True
presentation = ppt.Presentations.Add()
presentation.Slides.Add(1,1)
slides = presentation.Slides
slide1 = slides.Item(1)
shapes = slide1.Shapes

#%% 在ppt中创建一个excel图表
shape = shapes.AddOLEObject(ClassName='Excel.Chart')
pptWorkBook = shape.OLEFormat.Object
pptSheets = pptWorkBook.Sheets
pptSheet = pptSheets.Item('sheet1')
pptCharts = pptWorkBook.Charts
pptChart = pptCharts.Item(1)
seriesCollection = pptChart.SeriesCollection()
for series in seriesCollection :
    series.Delete()
