# -*- coding: utf-8 -*-
"""
Created on Thu Mar 06 11:22:03 2014

@author: Administrator
"""
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

#%% 打开一个测试文件
application = win32com.client.Dispatch('Excel.Application')
application.Visible = True
workbook = application.Workbooks.Open(r'c:\sample1.xls')
sheets = workbook.Sheets 
sheet1 = sheets.Item(1)

#%% 获取数据范围
nRow = sheet1.UsedRange.Rows.Count  
nCol = sheet1.UsedRange.Columns.Count
rangeName = cellName(1,1) + ':' + cellName(nRow,nCol)
rangeName

#%% 创建图表
chart = workbook.Charts.Add()
chart.ChartType = xlLine
chart.SetSourceData(sheet1.Range(rangeName))

#%% 图表类型说明
ChartType = {'xlLine':'折线图',
'xlLineMarkersStacked':'堆积数据点折线图',
'xlLineStacked':'堆积折线图',
'xlPie':'饼图',
'xlPieOfPie':'复合饼图',
'xlPyramidBarStacked':'堆积条形棱锥图',
'xlPyramidCol':'三维柱形棱锥图',
'xlPyramidColClustered':'簇状柱形棱锥图',
'xlPyramidColStacked':'堆积柱形棱锥图',
'xlPyramidColStacked100':'百分比堆积柱形棱锥图',
'xlRadar':'雷达图',
'xlRadarFilled':'填充雷达图',
'xlRadarMarkers':'数据点雷达图',
'xlStockHLC':'盘高-盘低-收盘图',
'xlStockOHLC':'开盘-盘高-盘低-收盘图',
'xlStockVHLC':'成交量-盘高-盘低-收盘图',
'xlStockVOHLC':'成交量-开盘-盘高-盘低-收盘图',
'xlSurface':'三维曲面图',
'xlSurfaceTopView':'曲面图（俯视图）',
'xlSurfaceTopViewWireframe':'曲面图（俯视框架图）',
'xlSurfaceWireframe':'三维曲面图（框架图）',
'xlXYScatter':'散点图',
'xlXYScatterLines':'折线散点图',
'xlXYScatterLinesNoMarkers':'无数据点折线散点图',
'xlXYScatterSmooth':'平滑线散点图',
'xlXYScatterSmoothNoMarkers':'无数据点平滑线散点图',
'xl3DArea':'三维面积图',
'xl3DAreaStacked':'三维堆积面积图',
'xl3DAreaStacked100':'百分比堆积面积图',
'xl3DBarClustered':'三维簇状条形图',
'xl3DBarStacked':'三维堆积条形图',
'xl3DBarStacked100':'三维百分比堆积条形图',
'xl3DColumn':'三维柱形图',
'xl3DColumnClustered':'三维簇状柱形图',
'xl3DColumnStacked':'三维堆积柱形图',
'xl3DColumnStacked100':'三维百分比堆积柱形图',
'xl3DLine':'三维折线图',
'xl3DPie':'三维饼图',
'xl3DPieExploded':'分离型三维饼图',
'xlArea':'面积图',
'xlAreaStacked':'堆积面积图',
'xlAreaStacked100':'百分比堆积面积图',
'xlBarClustered':'簇状条形图',
'xlBarOfPie':'复合条饼图',
'xlBarStacked':'堆积条形图',
'xlBarStacked100':'百分比堆积条形图',
'xlBubble':'气泡图',
'xlBubble3DEffect':'三维气泡图',
'xlColumnClustered':'簇状柱形图',
'xlColumnStacked':'堆积柱形图',
'xlColumnStacked100':'百分比堆积柱形图',
'xlConeBarClustered':'簇状条形圆锥图',
'xlConeBarStacked':'堆积条形圆锥图',
'xlConeBarStacked100':'百分比堆积条形圆锥图',
'xlConeCol':'三维柱形圆锥图',
'xlConeColClustered':'簇状柱形圆锥图',
'xlConeColStacked':'堆积柱形圆锥图',
'xlConeColStacked100':'百分比堆积柱形圆锥图',
'xlCylinderBarClustered':'簇状条形圆柱图',
'xlCylinderBarStacked':'堆积条形圆柱图',
'xlCylinderBarStacked100':'百分比堆积条形圆柱图',
'xlCylinderCol':'三维柱形圆柱图',
'xlCylinderColClustered':'簇状柱形圆锥图',
'xlCylinderColStacked':'堆积柱形圆锥图',
'xlCylinderColStacked100':'百分比堆积柱形圆柱图',
'xlDoughnut':'圆环图',
'xlDoughnutExploded':'分离型圆环图',
'xlLineMarkers':'数据点折线图',
'xlLineMarkersStacked100':'百分比堆积数据点折线图',
'xlLineStacked100':'百分比堆积折线图',
'xlPieExploded':'分离型饼图',
'xlPyramidBarClustered':'簇状条形棱锥图',
'xlPyramidBarStacked100':'百分比堆积条形棱锥图'}
