# -*- coding: utf-8 -*-
""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
                       变量作用域   
""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

#%% 可以是直接访问global数据
x = 1
def fun():
    print(x)           # 读取一个值 
fun()

#%% 可以是直接访问global数据
x = 1
def fun():
    y = x + 1          # 读取一个值 
    print 'y =', y      
fun()

#%% 
x = 1
def fun():
    x = 2                # 无法使用x=2，修改全局变量         
fun()
print "x =", x

#%%
x = 1
def fun():
    global x
    x = 2                # 指定了global x，所以x=2，可以修改全局变量
fun()
print 'x =' ,x

#%% 
x = 1
def fun():
    #import ch17-01      # 由于那个该死的"-"导致无法使用import语句导入模块 
    import sys
    sys.modules['ch17-01']

