# -*- coding: utf-8 -*-
"""
Created on Sun Jun 15 10:01:44 2014

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

@decorator
def func(a,b):
    print a + b        

#%% 以上代码和此等效
func = decorator(func)     #  *****理解这个很重要!!!******

#%%
func(1,2)

#%%
'''
思考：python实际上的函数实际上可以视为一个类的对象，该对象有一个方法叫做__call__当函
数被调用时该方法被调用。当使用类来定义一个decorator时，相当于用当前的func函数构造一个
decorator对象，并将对象的实例值返回并赋值给func。由于我们在decorator中定义了__call__
方法，所以返回来的decorator实例也是一个可以被调用的对象，等效于函数，这个值又赋给了
func对象，新的func实际是一个decorator对象，当func再被调用的时候，__call__方法就被调
用了。

注意：使用类格式定义的decorator不能应用于类方法。
    
'''