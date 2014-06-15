# -*- coding: utf-8 -*-
"""
Created on Sun Jun 15 12:34:45 2014

@author: Administrator
"""

#%%
class decorator:
    def __init__(self,func):
        self.func = func
    def __call__(self,*args):
        print('decorator.__call__ is begin')        
        self.func(*args)
        print('decorator.__call__ is end')
        


#%%
class C:
    @decorator
    def method(self,a,b):
        print a + b   

#%%
c = C()
c.method(1,2)




'''
使用类方法定义的decorator无法应用于类方法上，主要考虑一些原因：
1、类方法的定义的本质就是一个wrapper。
2、@装饰器定义符在对类方法所采取的行为和其对普通函数是有区别的。

以上两点具体参考相关脚本范例文件:
03.01…….py
03.02…….py
'''