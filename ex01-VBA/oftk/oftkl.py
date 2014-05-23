# -*- coding: utf-8 -*-
"""
需要模板调用的底层office实用方法包
"""
#%%
from wmi import WMI
from pandas.core.index import Index,Int64Index
import numpy as np

#%%
def clearSheet(sheet):
    '''
        将一个Excel sheet 的内容清空
    '''
    nRow = sheet.UsedRange.Rows.Count  
    nCol = sheet.UsedRange.Columns.Count
    for r in range(1,nRow+1):
        for c in range(1,nCol+1):
            sheet.Cells(r,c).Value = None

#%%
def copySheet(sheet1,sheet2):
    '''
        把一个Excel sheet 内容复制到另一个sheet
        sheet1 源
        sheet2 目标
    '''
    nRow = sheet1.UsedRange.Rows.Count  
    nCol = sheet1.UsedRange.Columns.Count
    for r in range(1,nRow+1):
        for c in range(1,nCol+1):
            sheet2.Cells(r,c).Value = sheet1.Cells(r,c).Value 

#%%
def df2sheet(df,sheet):
    u'''
        把一个DataFrame的内容拷贝到一个Excel sheet 中去        
        问题:目前仅支持单级索引
             没有检查数据框中的数据是否是unicode
    '''    
    # 检查数据框的索引类型    
    if (type(df.index) not in [Index,Int64Index]) or \
        (type(df.columns) not in [Index,Int64Index]) :
        raise Exception(u'olny DataFrame with Index allow!')
    
    # 清空目标sheet
    clearSheet(sheet)    
    
    # 添加行索引
    for (i,idx) in enumerate(df.index):
        
        sheet.Cells(2+i,1).Value = \
            idx.encode('gbk') if type(idx) == unicode else unicode(idx)
    
    # 添加列索引(columns)
    for (i,col) in enumerate(df.columns):
        sheet.Cells(1,2+i).Value = \
            col.encode('gbk') if type(col) == unicode else unicode(col)
        
    # 添加数据
    for (i,idx) in enumerate(df.index):
        for (j,col) in enumerate(df.columns):
            value = unicode(df.ix[df.index[i],df.columns[j]])
            sheet.Cells(2+i,2+j).Value = None if value == 'nan' else value           
           
                
#%%               
   
def terminateProcess(processName):
    '''
        按名称搜索并终止一个进行
    '''
    for i in WMI().Win32_Process(caption=processName):
        i.Terminate()

#%%        
def forceExcelQuit():
    '''
        确保Excel进程关闭
    '''
    terminateProcess('EXCEL.EXE')

#%%
def forcePowerPointQuit():
    '''
        确保ppt进程关闭
    '''
    terminateProcess('POWERPNT.EXE')
    

#%%    
def getPath():
    import inspect
    print(inspect.stack()[1][1])