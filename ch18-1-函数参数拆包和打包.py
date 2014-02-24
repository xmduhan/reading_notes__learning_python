# -*- coding: utf-8 -*-
""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
                    函数参数中的*和**
""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
#%% 调用者可以用*和**,分别对列表和字典进行解包
def f(a,b,c):
    print(a,b,c)
#%%
f(*[1,2,3])
#%%
f(*[1,2,3,4]) # 列表元素数量要和函数一致
#%%
f(**{'b':2,'c':3,'a':1,})


#%% 用*实现可变参数
def f(*args):
    for i in args:
        print(i)
#%% 
f(1)
#%%
f(1,2)
#%%
f(1,2,3)
#%%
f(*[1,2,3])
#%% 
f()

#%% 用**实现可变参数
def f(**kargs):
    for i in kargs:
        print(i,kargs[i])

#%%
f(a=1)
#%%
f(a=1,b=2)
#%%
f(**{'a':1,'b':2,'c':3})
#%%
f()

#%% 组合各种参数方式
def f(a,*args,**kargs):
    print('a=',1)
    for i in args:
        print(i)
    for i in kargs:
        print(i,kargs[i])
#%%
f()
#%%
f(1)
#%%
f(1,2)
#%%
f(1,2,c=3)
#%%
f(1,2,c=3,4)      # error


#%%  固定格式的拆包方式    
def f(a,(b,c)):
    print(a,b,c)    
x=1
y=(2,3)
f(x,y)
