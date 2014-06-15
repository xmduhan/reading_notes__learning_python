# -*- coding: utf-8 -*-
"""
Created on Sun Jun 15 09:09:29 2014

@author: Administrator
"""

#%%  
def decorator(something):
    print('decorator is call')    
    def wapper(*kargs):
        print('wapper begin')        
        something(*kargs)
        print('wapper end')
        
    return wapper

    
#%% 为一个函数添加装饰器
@decorator
def func(a,b): 
    print a+b    
    
#%% 以上代码和此等效
func = decorator(func)     #  *****理解这个很重要!!!******
    
#%%
func(1,2)

