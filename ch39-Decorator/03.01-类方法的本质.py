# -*- coding: utf-8 -*-
"""
Created on Sun Jun 15 11:54:41 2014

@author: Administrator
"""



#%%
class A:
    def m1(self,a):
        print a 
        
        
#%%
a = A()
m1 = a.m1
m1(1)

#%%
a.m1 = '1'
m1(1)         