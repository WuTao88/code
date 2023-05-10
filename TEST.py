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


def ZD(bm=[],H=0.0,L=30.0,data=[],i=0,Hsx=None):
    """_summary_

    Args:
        bm (list, optional):水准点 [zhunghao,'SZ01',564.3]
        H (float, optional):与水准点高差 H=H水准点-H待测点
        L (float, optional):与水准点的距离
        data (list, optional): _description_. Defaults to [].
        i (int, optional):转点编号 从默认0开始
        Hsx (_type_, optional):视线高

    Returns:
        返回下一次测点的视线高[]
    """
    
    name=bm[1]
    Hsz=bm[2]
    hd=abs(H/L*30) 
    
    if i==0 and Hsx==None:
        if hd <4 and hd>3.5:
            sz=round(random.uniform(hd+0.7,hd+1.0) if H<0 else random.uniform(0.7,1.0),3)
            dest=round(random.uniform(hd+0.7,hd+1.0) if H>0 else random.uniform(0.7,1.0),3)
        elif hd>=4:
            hd=2.0
            sz=round(random.uniform(hd+0.7,hd+1.0) if H<0 else random.uniform(0.7,1.0),3)
            dest=round(random.uniform(hd+0.7,hd+1.0) if H>0 else random.uniform(0.7,1.0),3)         
        else:
            sz=round(random.uniform(hd+1.0,hd+1.6) if H<0 else random.uniform(1.0,1.6),3)
            dest=round(random.uniform(hd+1.0,hd+1.6) if H>0 else random.uniform(1.0,1.6),3)
            hd=2.0
        if  abs(H)<4.1 and abs(L)<30:
            Hsx=round(Hsz+sz,3)
            data.append([name,sz,'H:%s'%(Hsz),'',Hsx,'',Hsz,''])
            return (Hsx,dest,1)
        else:
            data.append([name,sz,'','','','',Hsz,''])
        Hsx=round(Hsz+sz,3)
    else:
        sz=0
        dest=0
        hd=2.0 if hd>4.0 else hd
        
    i=1 if i==0 else i
    print(sz,dest)
    while round(H+sz-dest,3)!=0:
        hs=round(random.uniform(0.7+hd,1.0+hd) if H<0 else random.uniform(0.7,1.0),3)
        qs=round(random.uniform(0.7,1.0) if H<0 else random.uniform(0.7+hd,1.0+hd),3)
        if abs(round(H+sz-dest,3))<=hd and i>=int(L//30):
            qs=round(hs+H+sz-dest,3)                
            if qs<0.4:
                qst=round(random.uniform(0.7,0.7+hd),3)
                d=round(qst-qs,3)
                hs=round(hs+d,3)
                qs=qst
            elif qs>4.8:
                qst=round(random.uniform(4.5-hd,4.5),3)
                d=round(qst-qs,3)
                hs=round(hs+d,3)
                qs=qst
            Hsx=round(Hsx+hs-qs,3)
            
            data.append(['ZD%d'%i,'后视：%s'%(hs),'视线高：%s'%(Hsx),'前视：%s'%(qs),Hsx,'较水准点高差：','',''])
            
            break
        else:
            
            data.append(['ZD%d'%i,'后视：%s'%(hs),'视线高：%s'%(round(Hsx-qs+hs,3)),'前视：%s'%(qs),'较水准点高差：','','',''])
            Hsx=round(Hsx+hs-qs,3)
            H=round(H+hs-qs,3)
            
            i=i+1                           
    return (Hsx,dest,i+1)


def newSZ(m,a,b):
    h=0
    for i in range(m):
        pc=random.randint(a,b)
        if True:
            pass
data=[]
dd=ZD([123,'sz',67.86],6,5,data)



for d in data:
    print(d)