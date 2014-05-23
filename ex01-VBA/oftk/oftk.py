# -*- coding: utf-8 -*-
"""
提供较为复杂office实用方法包，主要为最终脚本封装服务
"""
#%%

import win32com 

from pandas.io.excel import ExcelWriter
from ppt import FourChartPage,ThreeChartPage
from oftkl import forcePowerPointQuit,forceExcelQuit

def dfs2Excel(dfs,sheetnames,filename):
    '''
        将一些列DataFrame保存到Excel中
    '''    
    excelFile = ExcelWriter(filename)
    for (i,df) in enumerate(dfs) :
        df.to_excel(excelFile,sheetnames[i])
    excelFile.save()     


#%% 
def addChart2Excel(filename,template):
    '''
         为一份Excel是添加图表
    '''
    excel = win32com.client.Dispatch('Excel.Application')
    #excel.Visible = True
    excel.DisplayAlerts = False
    workbook = excel.Workbooks.Open(filename)
    sheets = workbook.Sheets 
    for (i,sheet) in enumerate(sheets) : 
        if type(template) == list :
            template[i].render(sheet)
        else :
            template.render(sheet)
    workbook.Save()  
    workbook.Close()
    excel.Quit()
    forceExcelQuit()

#%%
def crtPPT4pT1(filename, mainTitle, titles, dfs, templates):
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    #powerpoint.Visible = True
    powerpoint.DisplayAlerts = False
    presentation = powerpoint.Presentations.Add()
    slide = presentation.Slides.Add(1,1)
    FourChartPage.render(slide, mainTitle, titles, dfs, templates)
    presentation.SaveAs(filename)
    presentation.Close()
    powerpoint.Quit()  
    forcePowerPointQuit()    

        
                
def crtPPT3pT1(filename, mainTitle, titles, dfs, templates):
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    #powerpoint.Visible = True
    powerpoint.DisplayAlerts = False
    presentation = powerpoint.Presentations.Add()
    slide = presentation.Slides.Add(1,1)
    ThreeChartPage.render(slide, mainTitle, titles, dfs, templates)
    presentation.SaveAs(filename)
    presentation.Close()
    powerpoint.Quit()  
    forcePowerPointQuit()    
    
    
    
