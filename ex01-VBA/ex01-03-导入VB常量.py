# -*- coding: utf-8 -*-

'''
为Python增加Office的VB常量
1、找到 “...\Python27\Lib\site-packages\win32com\client\makepy.py”，并运行。
2、根据实际需要选择要增加的常量库。
“Microsoft Office 1x.0 Object Library”  
“Microsoft PowerPoint 1x.0 Object Library”
“Microsoft Excel 1x.0 Object Library”
“Microsoft Word 1x.0 Object Library”
3、在“...\Python27\Lib\site-packages\win32com\gen_py”中会生成对应的python文件
，对其进行改名例如“Microsoft Office 1x.0 Object Library” ， 可以改名为msof。
4、然后在Python程序中使用以下方法引用
from win32com.gen_py import msof

注意：生成python的VB常量库后，不仅是可以使用VB的常量，com对象的类型识别也变得可能，可以
通过dir展示其所有的方法，但是属性还是只能查VBA参考，但是有些通过下标[]访问对象的方法会失效，
如slide = presentation.Slides[1] 不能再使用

'''

#%% 导入相关库
import win32com.client
from win32com.gen_py import msof,mspp

#%% 常量实际储存在constants对象中的
mspp.constants.ppLayoutBlank

#%% 将常量发布到全局命名空间中去
g = globals()
for c in dir(msof.constants) : g[c] = getattr(msof.constants, c)
for c in dir(mspp.constants) : g[c] = getattr(mspp.constants, c)

#%% 这样就可以像在VB中一样使用常量了
application = win32com.client.Dispatch("PowerPoint.Application")
application.Visible = True
presentation = application.Presentations.Add()
presentation.Slides.Add(1,ppLayoutBlank)
presentation.Slides.Add(2,ppLayoutChartAndText)
presentation.Slides.Add(3,ppLayoutFourObjects)

    
    
