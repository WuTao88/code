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


def deal(zhung_Hs:list,path):
    DATA=[]
    data=[]
    Hsx=0.0
    i=0
    zh0=0.0
    dest=0.0
    bm=None
    for zh_H in zhung_Hs:
        zh=zh_H[0]
        Hd=zh_H[1]
        if bm!=SP.get_BM(zh):
            if data !=[]:
                bm2=SP.get_BM(zh0)
                BM=BM=[bm2[2],bm2[3]]
                L2=round(math.sqrt(math.pow(zh0-bm2[0],2)+math.pow(bm2[1],2)),3)
                H2= round(Hsx-dest-Hd,3)
                print("i",i,'Hsx',Hsx)
                back=SP.ZD(BM,H2,L2,data,i,Hsx)
                Hsx=back[0]
                Ep=random.randint(-1*i,1*i)
                data.append([f'{BM[0]}','','',round(Hsx-BM[1],3),'',round(BM[1]+Ep/1000,3),BM[1],Ep])
                DATA.append(data)
                data=[]
                i=0
                zh0=0
                Hsx=0
            bm=SP.get_BM(zh)
            BM=[bm[2],bm[3]]
            print(bm)
            L=round(math.sqrt(math.pow(zh-bm[0],2)+math.pow(bm[1],2)),3) if zh0==0.0 else round(zh-zh0,3)
            H=round(BM[1]-Hd,3) if zh0==0.0 else round(Hsx-dest-Hd,3)        
            print(H,"<-H,L->",L)
            # print(BM)
            ce=SP.ZD(BM,H,L,data,i,Hsx)
            Δh=random.randint(-30,30)
            print("Δh==",Δh)
            data.append([SP.mileageToStr(zh),"","",round(Hsx-(Hd+Δh/1000),3),round(Hd+Δh/1000,3),Hd,Δh])        
            dest=ce[1]
            i=ce[2]
            Hsx=ce[0]
            zh0=zh
        else:

            L=round(math.sqrt(math.pow(zh-bm[0],2)+math.pow(bm[1],2)),3) if zh0==0.0 else round(zh-zh0,3)
            H=round(BM[1]-Hd,3) if zh0==0.0 else round(Hsx-dest-Hd,3)
            print(H,"<-H,L->",L)
            print(bm)
            Δh=random.randint(-30,30)
            if H>4.5 or H<0.5 or L>90:
                ce=SP.ZD(BM,H,L,data,i,Hsx)
                data.append([SP.mileageToStr(zh),"","",round(ce[0]-(Hd+Δh/1000),3),round(Hd+Δh/1000,3),Hd,Δh])
                dest=ce[1]
                i=ce[2]
                Hsx=ce[0]
                zh0=zh
            else:
                data.append([SP.mileageToStr(zh),"","",round(Hsx-(Hd+Δh/1000),3),round(Hd+Δh/1000,3),Hd,Δh])
                zh0=zh
    H=Hsx-BM[1]-dest
    
    L=round(math.sqrt(math.pow(zh-bm[0],2)+math.pow(bm[1],2)),3)
    back=SP.ZD(BM,H,abs(L),data,i,Hsx) if abs(L)!=0 else SP.ZD(BM,H,abs(L+2),data,i,Hsx)
    Hsx=back[0]
    print(L)
    if data!=[]:
        Ep=random.randint(-1*i,1*i)
        data.append([f'{BM[0]}','','',round(Hsx-BM[1],3),'',round(BM[1]+Ep/1000,3),BM[1],Ep])
        DATA.append(data)
    for d in DATA:
        print(d)
        print("===============================")



if __name__=="__main__":
    deal([[27670,626.23],[27675,627.23],[27703,626.83],[29709,626.53]],'test')
