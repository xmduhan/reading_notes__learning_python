# -*- coding: utf-8 -*-
"""
Created on Sun Mar 09 11:20:53 2014

@author: 14F
"""

#%% 导入必要的库文件
from __future__ import division
import pandas as pd
import os,sys
import dbtools
from pandas import Series,DataFrame
from xlchart import ColumnStacked
from oftk import addChart2Excel,dfs2Excel,crtPPT4pT1

#%% 设置目录和文件名称
path = r'D:\工作记录\01-体系类\01-0.经营服务分析会\2014经营分析会报告\00-数据提取及页面生成\01-天翼离网分析'
path = unicode(path,'utf-8')
os.chdir(path)
sys.path.append(path)
filename = u'%s-01-天翼(分市场)离网量趋势'
excelFileName = (path +  u'\\' + filename +  u'.xls') % '2'
pptFileName = (path +  u'\\' + filename +  u'.ppt') % '3'

#%% 导入数据
con = dbtools.getConnection('sjck').connection
data = pd.io.sql.read_frame('select * from duh0309_1_ty_market',con)
data = data.applymap(lambda x: x.decode('gbk') if type(x) == str else x )

#%% 定义数据处理过程
def proc(d):
    d = d.groupby(['YEAR','MONTH_ID']).sum()
    thisYear = Series(zip(*d.index)[0]).max()
    lastYear = Series(zip(*d.index)[0]).min()
    d[u'正常离网'] = d.OUT - d.BACK - d.WX       # OUT-BACK = 离网
    d[u'无效离网'] = d.WX                        # WX 已经包含 OUT-BACK
    d[u'回网'] = -d.BACK
    d1_thisYear = d.ix[thisYear:][1:]
    d1_lastYear = d.ix[:lastYear][len(d1_thisYear)-13:]
    d = pd.concat([d1_thisYear,d1_lastYear])
    d = d.sort_index()
    d = d.reset_index()
    d[u'月份'] = d.YEAR.apply(str) + d.MONTH_ID.apply(lambda x:str(x).rjust(2,'0'))
    d=d[[u'月份',u'正常离网',u'无效离网',u'回网']]
    d = d.set_index(u'月份')
    d.index.name = None
    return d

#%% 生成全网数据 
d0 = proc(data)
d0

#%% 普通天翼
d1 = proc(data[data.MARKET==u'普通天翼'])
d1

#%% 校园套餐
d2 = proc(data[data.MARKET==u'校园套餐'])
d2

#%% 无线宽带
d3 = proc(data[data.MARKET==u'无线宽带'])
d3

#%%
dfs=[d0,d1,d2,d3]
labels = [u'全网天翼',u'普通天翼',u'校园套餐',u'无线宽带']

#%% 保存到excel文件, 并为excel添加图表
dfs2Excel(dfs,labels,excelFileName)
addChart2Excel(excelFileName,ColumnStacked)

#%% 生成图表的标题
mainTitle = u'天翼分市场离网量趋势'
nums = [u'①',u'②',u'③',u'④']
titles = []
for i in range(4) :
    df = dfs[i]
    label = labels[i]
    num = nums[i]
    out  = df.ix[-1,u'正常离网'] + df.ix[-1,u'无效离网']
    wx = df.ix[-1,u'无效离网']
    outLm = df.ix[-2,u'正常离网'] + df.ix[-2,u'无效离网']
    outLy = df.ix[0,u'正常离网'] + df.ix[0,u'无效离网']
    wordLm = u'增加' + unicode(out-outLm) if out >= outLm  else  u'减少' + unicode(outLm-out)
    wordLy = u'增加' + unicode(out-outLy) if out >= outLy  else  u'减少' + unicode(outLy-out)
    title = u'%s  %s：本月离网%d，其中无效客户%d；较上月%s部，较去年同期%s部。'
    title = title % (num,label,out,wx,wordLm,wordLy)
    titles.append(title)
    
#%% 制作ppt
crtPPT4pT1(pptFileName, mainTitle, titles, dfs, [ColumnStacked] * 4)

