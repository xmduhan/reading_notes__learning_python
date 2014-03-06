# -*- coding: utf-8 -*-
"""
Created on Thu Mar 06 09:59:50 2014

@author: Administrator
"""
#%%
import win32com.client
from win32com.gen_py import msof,mspp,msxl

#%% 打开excel程序
application = win32com.client.Dispatch('Excel.Application')
application.Visible = True

#%% 打开一个工作簿
workbook = application.Workbooks.Open(r'c:\sample1.xls')

#%% 遍历每一个工作表
sheets = workbook.Sheets 
for sheet in sheets :
    print sheet.Name

#%% 访问工作表中的数据
sheet1 = sheets.Item(1)
print sheet1.Cells(1,2)

#%% 侦测数据表使用范围
print sheet1.UsedRange.Rows.Count  
print sheet1.UsedRange.Columns.Count

#%% 修改数据
sheet1.Cells(1,2).Value = u'2013年'.encode('gbk')

#%% 修改数据背景颜色
sheet1.Cells(3,1).Interior.Color = 0xFFFF #by RGB  
sheet1.Cells(3,2).Interior.ColorIndex = 15 #by ColorIndex

#%% 修改字体和颜色
sheet1.Cells(4,3).Font.Bold = True  
sheet1.Cells(3,3).Font.Color = 0xFF 

#%% 保存数据退出
workbook.Save()  
workbook.Close()  
application.Quit()  