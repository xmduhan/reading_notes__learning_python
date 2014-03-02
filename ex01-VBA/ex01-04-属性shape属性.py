# -*- coding: utf-8 -*-
"""
Created on Sun Mar 02 00:00:41 2014

@author: 管理员
"""
#%% 导出模块和VB常量
import win32com.client
from win32com.gen_py import msof,mspp
g = globals()
for c in dir(msof.constants) : g[c] = getattr(msof.constants, c)
for c in dir(mspp.constants) : g[c] = getattr(mspp.constants, c)

#%% 创建shape类型的值和名称的对照
shapeTypeNameList=['msoAutoShape','msoCallout','msoCanvas','msoChart','msoComment',  
'msoDiagram','msoEmbeddedOLEObject','msoFormControl','msoFreeform',  
'msoGroup','msoLine','msoLinkedOLEObject','msoLinkedPicture', 'msoMedia',
'msoOLEControlObject','msoPicture','msoPlaceholder','msoScriptAnchor',
'msoShapeTypeMixed','msoTable','msoTextBox','msoTextEffect']
shapeType = {g[i]:i for i in shapeTypeNameList}
autoShapeType = { g[i]:i for i in g if i.startswith('msoShape')}

#%% 导入一个测试用的ppt
application = win32com.client.Dispatch("PowerPoint.Application")
application.Visible = True
filename = r'c:\sample1.ppt'
presentation = application.Presentations.Open(filename)
slides = presentation.Slides

#%% 打印出每一个对象(shape)的属性
for (i,slide) in enumerate(slides):
    print 'page' + str(i) + ':'
    for (j,shape) in enumerate(slide.Shapes):
        print ' shape%d:%s,%s,(%f,%f)(%f,%f)' % (j,shapeType[shape.Type],\
            autoShapeType[shape.AutoShapeType],shape.Left,shape.Top,\
            shape.Width,shape.Height)

#%% 对属性进行修改            
for (i,slide) in enumerate(slides):
    for (j,shape) in enumerate(slide.Shapes):
        shape.Left = 0
        shape.Top = 0