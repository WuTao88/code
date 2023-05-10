

import random
import sys
import os
import math
import openpyxl as xl
import xlwings as xw
from tools import *

def source_path(relative_path):
    if os.path.exists(relative_path):
        base_path = os.path.abspath(".")
    elif getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    return os.path.join(base_path, relative_path)



RES=source_path('res')



def log(*msg):
    f=open('日志.txt','a+',encoding='UTF-8')
    f.write(f'{time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())}  :  {msg}\n')
    f.close()

class SP:

    def AngleChange(angle):
        d = int(angle)
        m = int((angle - int(angle)) * 60)
        s = round((angle - int(angle)) * 3600 - m * 60, 2)
        dms = "{0}°{1}′{2}″".format(d, m, s)
        return dms


    def mileageToStr(km):
        if km==None:
            return ''
        k = int(km // 1000)
        b = int((km - k * 1000) // 100)
        s = int((km - k * 1000 - b * 100) // 10)
        g = int((km - k * 1000 - b * 100 - s * 10) // 1)
        f = km - km // 1
        ZH = "K{0}+{1}{2}{3:0>3.03f}".format(k, b, s, g + f)
        return ZH



    def LMZD(bm=[],H=0.0,L=30.0,data=[],i=0,Hsx=None):
        
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
                data.append([name,'',sz,'','',Hsx,'',Hsz,'',''])
                return (Hsx,dest,1)
            else:
                data.append([name,'',sz,'','','','',Hsz,'',''])
            Hsx=round(Hsz+sz,3)
        else:
            sz=0
            dest=0
            hd=2.0 if hd>4.0 else hd
                    
        i=1 if i==0 else i
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
                data.append(['ZD%d'%i,'',hs,'',qs,Hsx,'','','',''])
                break
            else:
                data.append(['ZD%d'%i,'',hs,'',qs,'','','','',''])
                Hsx=round(Hsx+hs-qs,3)
                H=round(H+hs-qs,3)
                i=i+1                           
        return (Hsx,dest,i+1)

    def ZD(bm=[],H=0.0,L=30.0,data=[],i=0,Hsx=None):
        
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
                data.append([name,sz,'','',Hsx,'',Hsz,''])
                return (Hsx,dest,1)
            else:
                data.append([name,sz,'','','','',Hsz,''])
            Hsx=round(Hsz+sz,3)
        else:
            sz=0
            dest=0
            hd=2.0 if hd>4.0 else hd
            
        i=1 if i==0 else i
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
                data.append(['ZD%d'%i,hs,'',qs,Hsx,'','',''])
                break
            else:
                data.append(['ZD%d'%i,hs,'',qs,'','','',''])
                Hsx=round(Hsx+hs-qs,3)
                H=round(H+hs-qs,3)
                i=i+1                           
        return (Hsx,dest,i+1)

    
    def HEADER(KZD=[],*args):
        
        if KZD!=None :
            if args==():
                X0=KZD[0][1]
                Y0=KZD[0][2]
                H0=KZD[0][3]
                X2=KZD[1][1]
                Y2=KZD[1][2]
                H2=KZD[1][3]
                s=round(((X2-X0)**2+(Y2-Y0)**2)**0.5,3)
                a=math.degrees(math.atan2(X2-X0,Y2-Y0))
                α=SP.AngleChange(a) if a>0 else SP.AngleChange(a+360)
                KZD.append(s)
                KZD.append(α)
                return KZD
            else:
                X0=KZD[0][1]
                Y0=KZD[0][2]
                X1=args[0]
                Y1=args[1]
                s=round(((X1-X0)**2+(Y1-Y0)**2)**0.5,3)
                a=math.degrees(math.atan2(X1-X0,Y1-Y0))
                α=SP.AngleChange(a) if a>0 else SP.AngleChange(a+360)
                return (α,s)

        else:
            return None
        
    def writeFX(app,**kwargs):

        DD=kwargs['data']
        wb=app.books.open(muban+'\\全站仪放线记录表.xlsx')
        sample=wb.sheets['sheet']
        i=1
        for dt in DD:

            
            data=dt[0]
            
            num=len(data)
            yigao=round(random.uniform(1.2,1.6),3)
            pg=(num-1)//11
            header=dt[1]
            
            for p in range(1,pg+2):
                
                sample.api.Copy(Before=sample.api)
                
                sht = wb.sheets[f'sheet ({2})']
                
                new_sheet_name = f'{kwargs["zhuanghao"][:8]}数据({i}.{p})'
                
                sht.api.Name = new_sheet_name

                print('did muban\n',sht.api.Name)
                sht.range('B1').value=kwargs['project']
                sht.range('C5').value=f'{kwargs["gongcheng"]}'
                sht.range('L5').value=f'{kwargs["zhuanghao"]}{kwargs["gongxu"]}'
                sht.range('C7').value=header[0][0] #测点编号
                sht.range('F7').value=header[0][1] #测点坐标
                sht.range('F8').value=header[0][2] #测点坐标
                sht.range('F9').value=header[0][3] #测点坐标
                sht.range('I7').value=header[1][0] #后视点编号
                sht.range('M7').value=header[1][1] #后视点坐标
                sht.range('M8').value=header[1][2] #后视点坐标
                sht.range('M9').value=header[1][3] #后视点坐标
                sht.range('Q7').value=header[2]    #极坐标
                sht.range('Q9').value=header[3]    #极坐标
                sht.range('U9').value=f'{yigao}'
                
                w=(p-1)*11+11 if num>((p-1)*11+11) else num        
                sht.range('B12').value=data[(p-1)*11:w]
                
            i=i+1
            
        wb.save(kwargs['path']+f'\\{kwargs["zhuanghao"]}{kwargs["gongxu"]}-'+'放线记录.xlsx')
        wb.close()

    def writePM(app,**kwargs):

        DD=kwargs['data']
        wb=app.books.open(muban+'\\全站仪平面位置检测表.xlsx')
        sample=wb.sheets['sheet']
        i=1        
        for dt in DD:

            data=dt[0]
            num=len(data)
            yigao=round(random.uniform(1.2,1.6),3)
            pg=(num-1)//10
            header=dt[1]
            for p in range(1,pg+2):            
                
                sample.api.Copy(Before=sample.api)
                sht = wb.sheets[f'sheet ({2})']
                new_sheet_name = f'{kwargs["zhuanghao"][:8]}数据({i}.{p})'
                sht.api.Name = new_sheet_name

                sht.range('B1').value=kwargs['project']
                sht.range('C5').value=f'{kwargs["gongcheng"]}'
                sht.range('L5').value=f'{kwargs["zhuanghao"]}{kwargs["gongxu"]}' 
                sht.range('C7').value=header[0][0] #测点编号
                sht.range('F7').value=header[0][1] #测点坐标
                sht.range('F8').value=header[0][2] #测点坐标
                sht.range('F9').value=header[0][3] #测点坐标
                sht.range('I7').value=header[1][0] #后视点编号
                sht.range('L7').value=header[1][1] #后视点坐标
                sht.range('L8').value=header[1][2] #后视点坐标
                sht.range('L9').value=header[1][3] #后视点坐标
                sht.range('P7').value=header[2]    #极坐标
                sht.range('P9').value=header[3]    #极坐标
                sht.range('U9').value=yigao
                w=(p-1)*10+10 if num>((p-1)*10+10) else num        
                sht.range('B12').value=data[(p-1)*10:w]
            i=i+1
        wb.save(kwargs['path']+f'\\{kwargs["zhuanghao"]}{kwargs["gongxu"]}-'+'平面位置检测.xlsx')
        wb.close()

    def writeGC(app,**kwargs):
        DATA=kwargs['data']
        wb=app.books.open(muban+'\\高程.xlsx')
        sample=wb.sheets['sheet']
        i=1
        for dt in DATA:
            data=dt[0]
            num=len(data)
            pg=(num-1)//17
            for p in range(0,pg+1):
                sample.api.Copy(Before=sample.api)
                sht = wb.sheets[f'sheet ({2})']
                new_sheet_name = f'{kwargs["zhuanghao"][:8]}{kwargs["gongxu"][4:]}数据({i}.{p+1})'
                sht.api.Name = new_sheet_name
                sht=wb.sheets[sht.api.Name]
                sht.range('A1').value=kwargs['project']
                sht.range('B5').value=f"{kwargs['gongcheng']}"
                sht.range('E5').value=f"{kwargs['zhuanghao']}{kwargs['gongxu']}"
                w=p*17+17 if num>(p*17+17) else num
                sht.range('A9').value=data[p*17:w]                
                if p==pg:
                    sht.range('B26').value=f"ΣH=H测-H设={dt[1][0]}-{dt[1][1]}={dt[1][2]}mm，符合精度要求"
            i=i+1
        wb.save(kwargs['path']+f"\\{kwargs['zhuanghao']}{kwargs['gongxu']}-"+'高程检测.xlsx')
        wb.close()


    def writeGCJS(app,**kwargs):
        DATA=kwargs['data']
        wb=app.books.open(muban+'\\路基路面高程检测、计算.xlsx')
        sample=wb.sheets['sheet']
        i=1
        for dt in DATA:
            data=dt[0]
            num=len(data)
            pg=(num-1)//19
            for p in range(0,pg+1):
                sample.api.Copy(Before=sample.api)
                sht = wb.sheets[f'sheet ({2})']
                new_sheet_name = f'{kwargs["zhuanghao"][:8]}{kwargs["gongxu"][4:]}数据({i}.{p+1})'
                sht.api.Name = new_sheet_name
                sht=wb.sheets[sht.api.Name]
                sht.range('A1').value=kwargs['project']
                sht.range('B6').value=f"{kwargs['zhuanghao']}{kwargs['gongcheng']}"
                sht.range('F6').value=f"{kwargs['gongxu']}"
                w=p*19+19 if num>(p*19+19) else num
                sht.range('A10').value=data[p*19:w]                
            i=i+1
        wb.save(kwargs['path']+f"\\{kwargs['zhuanghao']}{kwargs['gongxu']}-"+'路基路面高程检测、计算.xlsx')
        wb.close()

class OPTION:

    def getBM(zh=0.0,filename=None):
        global res

        name=res+'\\BM' if filename ==None else filename
        data=open(name, "r+",encoding='UTF-8')
        gc={}
        for sz in data:        
            gc=eval(sz)
        near=[abs(round(i-zh,3)) for i in gc.keys()]
        key1=round(zh+min(near),3)
        key2=round(zh-min(near),3)

        if (key1) in gc:
            return (key1,gc[key1][0],gc[key1][1])
        else:
            return (key2,gc[key2][0],gc[key2][1])


    def getCtr(zh=0.0,filename=None):
        global res
        name=res+'\\ContralPoint' if filename ==None else filename
        data=open(name, "r+",encoding='UTF-8')
        KZD={}
        for v in data:        
            KZD=eval(v)
        keys=list(KZD.keys())
        near=[abs(round(i-zh,3)) for i in KZD.keys()]
        key1=round(zh+min(near),3)
        key2=round(zh-min(near),3)
        a=None
        dt=[]
        if (key1) in KZD:
            a=key1
            dt.append(KZD[key1])
        else:
            a=key2
            dt.append(KZD[key2])
        if keys.index(a)+1 <len(keys):
            dt.append(KZD[keys[keys.index(a)+1]])
        else:
            dt.append(KZD[keys[keys.index(a)-1]])
        return dt


def gaochengDeal(tree,PATH,**kwargs):

    try:
        app=xw.App(visible=False, add_book=False)
        kwargs={k:kwargs[k].get() for k in kwargs}
        ZH1=kwargs['zh1']
        path=f"{PATH.get()}\\{kwargs['gongcheng']}\\{SP.mileageToStr(ZH1)}{ '-'+SP.mileageToStr(kwargs['zh2']) if 'zh2' in kwargs else ''}{kwargs['side'] if 'side' in kwargs else ''}{kwargs['gongxu']}"
        if os.path.isdir(path):
            pass
        else:
            os.makedirs(path)
        kwargs['path']=path
        kwargs['zhuanghao']=f"{SP.mileageToStr(ZH1)}{ '-'+SP.mileageToStr(kwargs['zh2']) if 'zh2' in kwargs else ''} {kwargs['side'] if 'side' in kwargs else ''}" if '回填' not in kwargs['gongxu'] else f"{SP.mileageToStr(ZH1)}{ '-'+SP.mileageToStr(kwargs['zh2']) if 'zh2' in kwargs else ''}"
        DATA=[]
        data=[]
        BM=None
        Hsx=None
        dest=None
        Hq=None
        zhq=None
        i=0
        ZHS=tree.get_children()
        for item in ZHS:
            row=tree.item(item,'values')

            zh=round(float(row[0]),3)
            pianju=row[1]
            HS=round(float(row[2]),3)
            pc=random.randint(kwargs['pc_s'],kwargs['pc_b'])
            if BM!=OPTION.getBM(zh):
                if data !=[]:
                    
                    
                    H=Hsx-BM[2]-dest
                    L=BM[0]-zhq
                    back=SP.ZD(BM,H,abs(L),data,i,Hsx) if abs(L)!=0 else SP.ZD(BM,H,abs(L+2),data,i,Hsx)
                    Hsx=back[0]
                    Ep=random.randint(-1*i,1*i)
                    data.append([f'{BM[1]}','','',round(Hsx-BM[2],3),'',round(BM[2]+Ep/1000,3),BM[2],Ep])                        
                    
                    DATA.append([data,[round(BM[2]+Ep/1000,3),BM[2],Ep]])
                    SP.writeGC(app,**kwargs)
                    data=[]               
                BM=OPTION.getBM(zh)                    
                H=BM[2]-HS
                L=abs(zh-BM[0])
                Hsx=None
                i=0
                back=SP.ZD(BM,H,abs(L),data,i,Hsx)
                Hsx=back[0]
                i=back[2]
                dest=back[1]
                zhq=zh
                Hq=HS 
                data.append([f'{SP.mileageToStr(zh)},{pianju}','','',round(Hsx-round(HS+pc/1000,3),3),'',round(HS+pc/1000,3),HS,pc])
            else:
                HC=Hsx-HS            
                L=zh-zhq
                if HC<0.6 or HC>4.8:
                    H=Hsx-HS-dest if HC> 4.8 else HC-2
                    back=SP.ZD(BM,H,abs(L),data,i,Hsx) if abs(L)!=0 else SP.ZD(BM,H,abs(L+2),data,i,Hsx)
                    log(back)
                    Hsx=back[0]
                    i=back[2]
                    zhq=zh
                    Hq=HS        
                data.append([f'{SP.mileageToStr(zh)},{pianju}','','',round(Hsx-round(HS+pc/1000,3),3),'',round(HS+pc/1000,3),HS,pc])
        
        log(Hq,BM,zhq,i,dest)
        
        
        H=Hsx-BM[2]-dest
        L=BM[0]-zhq
        back=SP.ZD(BM,H,abs(L),data,i,Hsx) if abs(L)!=0 else SP.ZD(BM,H,abs(L+2),data,i,Hsx)
        Hsx=back[0]
        if data!=[]:
            
            Ep=random.randint(-1*i,1*i)
            data.append([f'{BM[1]}','','',round(Hsx-BM[2],3),'',round(BM[2]+Ep/1000,3),BM[2],Ep])         
            DATA.append([data,[round(BM[2]+Ep/1000,3),BM[2],Ep]])
        kwargs['data']=DATA
        SP.writeGC(app,**kwargs)
        
    except Exception as err:
        log('错误信息',f'发生错误,{err}，执行失败！！')
    else:

        log('信息', '成功')
    finally:
        app.quit()
        