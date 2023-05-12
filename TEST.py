#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :    TEST.py
@Time    :    2023/05/08 20:17:34
@Author  :    cyq
@Version :    1.0
@Contact :    1135362921@qq.com
@Desc    :    
'''


from shuizhun import *
import openpyxl as xl
import xlwings as xw



def office(args,**kwargs):


    pythoncom.CoInitialize()
    app=xw.App(visible=False, add_book=False)
    SP.writeGC(app,**kwargs)




if __name__=="__main__":
    data=[]
    dd=SP.ZD(['sz',564.884],85.846,5,data)
    
    DATA=data
    path='C:\\Users\\Tao\\Desktop\\test'
    kw={'path':path,'data':data,'zhuanghao':'33600-33760','gongxu':'边沟','gongcheng':'test','side':'右侧'}
    office('None',**kw)
    print(dd)
