# -*- coding: utf-8 -*-
"""
Created on Sun Oct 02 12:29:12 2016

@author: Bassant
"""

import cv2
import numpy as np
from reportlab.lib.pagesizes import  A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm

import string
import xlwt
import time


global cen,sh2,AnsRef,SBnRef,Anss,CorAns,CorAnsPos,AvAns, sheet,sheetTh,sheetTh2,sheetTh3,minus_G,neg_G,NAv
global Results,Q,QpC,cols,colsr,case,Ques_Stat,Ans_Stat,CAth,StuSheet,StuSheet2,Stu_Cor_Ans
global CorQ,wQ,MisQ,grade,nSec,stpsec,SCth,Stu_Qs_State   
       
ColRef={}
ColRefTh={}
cen=[]
AnsRef={}
SBnRef={}
sheet={}
sheetTh={}
sheetTh2={}
sheetTh3={}
Results={}
Ques_Stat={}
Ans_Stat={}
CorAns={}
CorAnsPos={}
CAth={} 
StuSheet={}
StuSheet2={}
Stu_Cor_Ans={}
SCth={}
Stu_Qs_State=[]
CorQ={}  #correct questions
WQ={}
MisQ={}
Q=[]

cols=["col 0","col 1","col 2","col 3"]
cols.reverse()
colsr=["col 0", "col 1", "col 2" , "col 3"]
case=["C","W","M"]

 

def Ans_Sheet(self,*args):
    start_time = time.time()
    #print "start " ,start_time
    
    NQ=int(self.ids.nq.text)
    NAv=int(self.ids.nc.text)
    title=str(self.ids.ttl.text)
    cls=str(self.ids.clss.text)
    date=str(self.ids.dt.text)
    
    AvAns=string.ascii_uppercase[0:NAv]
    
    c=canvas.Canvas(title+".pdf", pagesize=A4)
    c.setFont("Helvetica", 8)
    c.setStrokeColorRGB(0,0,0)
    c.setFillColorRGB(0.2,0.2,0.2)
    
    c.drawString(14*cm,26.9*cm+(0.6*2)*cm,  title)

    c.drawString(14*cm,26.9*cm+0.6*cm,  cls)

    c.drawString(14*cm,26.9*cm,  date)


    c.line(0.2*cm,26*cm,20.8*cm,26*cm)
    c.line(0.2*cm,26.1*cm,20.8*cm,26.1*cm)


    QpC= int(np.ceil(NQ/4.0))
    Q=[QpC,QpC,QpC,NQ-3*QpC]
    
    if Q[3] > Q[0]:
        btlin=Q[3]
    else:
        btlin=Q[0]
    
    for col in range(4):
        for q in range(Q[col]):
            c.drawString((1.2*cm)+(5*col)*cm ,(25-0.6*q)*cm, "Q"+ str(q+1+QpC*col))
            for i in range(NAv):
                x=2.4*cm+(i*0.7)*cm+(5*col)*cm
                y=(25-0.6*q)*cm
                r=0.2*cm
                c.ellipse(x-r, y-r, x+r, y+r, stroke=1,fill=0)
                c.drawString(x-0.08*cm,y-0.1*cm, AvAns[i])

        c.line(x+0.4*cm,25.6*cm,x+0.4*cm ,(25.2-0.61*btlin)*cm) #after col
    
        c.line(x+0.15*cm,25.35*cm,x+0.65*cm ,25.35*cm)  # -
    
        c.line(x+0.15*cm,(25.4-0.61*btlin)*cm,x+0.65*cm ,(25.4-0.61*btlin)*cm)  # -
    
    c.line(0.5*cm,25.6*cm,0.5*cm ,(25.2-0.61*btlin)*cm) #before 1st col
    c.line(0.25*cm,25.35*cm,0.75*cm ,25.35*cm)  # -
    c.line(0.25*cm,(25.4-0.61*btlin)*cm,0.75*cm ,(25.4-0.61*btlin)*cm)  # -

    c.line(0.25*cm,(24.9-0.6*btlin-0.3)*cm,20.8*cm,(24.9-0.6*btlin-0.3)*cm) #bottom

    l=[[2.3*cm,26.5*cm,9.1*cm,26.5*cm], [2.3*cm, 28.5*cm, 9.1*cm, 28.5*cm] ]
    c.lines(l)
    l2=[[2.55*cm,26.25*cm,2.55*cm,26.75*cm], [2.55*cm,28.25*cm,2.55*cm,28.75*cm], 
        [8.85*cm, 26.25*cm,8.85*cm, 26.75*cm], [8.85*cm, 28.25*cm, 8.85*cm, 28.75*cm] ]
    c.lines(l2)
    inf=["BN-U","BN-T","Sec"]        
    for d in range(3):
        c.drawString(2*cm-0.1*cm,26.9*cm+(0.6*d)*cm-0.1*cm,  inf[d])
        for i in range(10):
            x= 3*cm+(i*0.6)*cm
            y=26.9*cm+(d*0.6)*cm
            r=0.2*cm
            c.ellipse(x-r, y-r, x+r, y+r, stroke=1,fill=0)
            c.drawString(x-0.07*cm,y-0.1*cm, str(i))    
    c.showPage()
    c.save()
    
    #print("--- %s seconds ---" % (time.time() - start_time))
''' -------------------------------------------------------------------'''    
    #------------------------------------------------------------------------#
def SheetInfo(self,*args):
    
    minus_G=False
    NQ=int(self.ids.nq.text)
    NAv=int(self.ids.nc.text)
    CAG=float(self.ids.cag.text)
    neg_G=0
    if (self.ids.cwn.state)== 'down': 
        minus_G=True 
        neg_G=float(self.ids.ng.text)
    ONT=float(self.ids.ont.text)
    nSec=int(self.ids.nsec.text)
    stpsec=int(self.ids.stu.text)
    title=str(self.ids.ttls.text)
    EMT=str(self.ids.emt.text)
    SF=str(self.ids.sf.text)
    stsheets=str(self.ids.sf.text)
    emptysheet=cv2.imread(self.ids.emt.text)
    
    
    AvAns=string.ascii_uppercase[0:NAv]
    QpC= int(np.ceil(NQ/4.0))
    Q=[QpC,QpC,QpC,NQ-3*QpC]
    Q.reverse()
    
    return NQ,NAv,CAG,minus_G,neg_G,ONT,nSec,stpsec,title,EMT,SF,AvAns,Q,QpC,emptysheet,stsheets
''' -------------------------------------------------------------'''
def RDic(nSec,stpsec):
    for n in range(nSec):
        ns=str(n+1)
        Results[ns]=[]
        for s in range(stpsec):
            Results[ns].append(0)   
    print "Results= " ,Results
    return Results
''' -----------------------------------------------------------------'''
    
def EmptySheet(img):
    
    gray= cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
    un, imth=cv2.threshold(gray,127,255,cv2.THRESH_BINARY_INV)  #160
    #cv2.imshow("EmptySheet_TH", imth)
    cv2.imwrite("EmptySheet_TH.png", imth)

    imfloodfill = imth.copy()
    h, w = imth.shape[:2]
    mask = np.zeros((h+2, w+2), np.uint8)
    cv2.floodFill(imfloodfill, mask, (0,0), 255)
    #cv2.imshow("EmptySheet_FF",imfloodfill)
    cv2.imwrite("EmptySheet_FF.png",imfloodfill)
    
    imfloodfillinv = cv2.bitwise_not(imfloodfill)
    #cv2.imshow("EmptySheet_FFI",imfloodfillinv)
    cv2.imwrite("EmptySheet_FFI.png",imfloodfillinv)
    
    imout = imth | imfloodfillinv
    #cv2.imshow("EmptySheet_filled", imout)
    cv2.imwrite("EmptySheet_filled.png", imout)
    
    imout = cv2.medianBlur(imout,9)
    #cv2.imshow("EmptySheet_filled_blured",imout)
    cv2.imwrite("EmptySheet_filled_blured.png",imout)
    
    imout2=imout.copy()
    #cv2.imshow("th2",imout2)
    imout3=imout.copy()
    
    corners = cv2.goodFeaturesToTrack(gray,14,0.9,5) #0.4
    corners = np.int0(corners)

    crn=[]
    for i in corners:
       t= tuple(i[0])
       crn.append(t)
    crn.sort()
    cen=[]
    for i in range(14):
        cen.append(tuple(corners[i][0]))    
    cen=crn
    
    SecBn=[cen[2],cen[3],cen[6],cen[7]]
    col0=[cen[0],cen[1],cen[4],cen[5]]
    col1=[cen[4],cen[5],cen[8],cen[9]]
    col2=[cen[8],cen[9],cen[10],cen[11]]
    col3=[cen[10],cen[11],cen[12],cen[13]]
    
    '''
    SecBn=[cen[2],cen[3],cen[6],cen[7]]
    col0=[cen[0],cen[1],cen[4],cen[5]]
    col1=[cen[4],cen[5],cen[8],cen[9]]
    col2=[cen[8],cen[9],cen[10],cen[11]]
    col3=[cen[10],cen[11],cen[12],cen[13]]
    '''
    cen={"SecBn":SecBn,"col 0":col0,"col 1":col1,"col 2":col2,"col 3":col3}

    for i in cen:  
       y1=int(cen[i][0][1])
       y2=int(cen[i][1][1])
       x1=int(cen[i][0][0])
       x2=int(cen[i][2][0])
       col=img[y1:y2,x1:x2]       
       sheet[i]=col
       col=imout[y1:y2,x1:x2]
       sheetTh[i]=col
       col2=imout2[y1:y2,x1:x2]
       sheetTh2[i]=col2
       col3=imout3[y1:y2,x1:x2]
       sheetTh3[i]=col3
    
       #cv2.imshow("EmptySheet-"+i, sheetTh[i]) 
       cv2.imwrite("EmptySheet-TH-"+i+".png", sheetTh[i])
       cv2.imwrite("EmptySheet-"+i+".png", sheet[i])
       
    return cen, sheet, sheetTh,sheetTh2,sheetTh3

'''' ----------------------------------------------------------------'''
    
''' ----------------------------------------------------------------------- '''
   
def Extract_Pos(sheetTh,NAv,AvAns):
    global r
    r=0
    
    for i in range(4):
        ans={}
        for ci in AvAns:
            ans[ci]=[]
            
        contours, hierarchy = cv2.findContours (sheetTh["col "+str(i)],cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
        contours.reverse()
        
        for k in range(NAv):
           temp=[]
           # to compare other centers to make sure they're the same question
           (x,y),radiusf = cv2.minEnclosingCircle(contours[k])
           center = (int(x),int(y))
           for j in range(len(contours)):    
                 (xj,yj),radiusj = cv2.minEnclosingCircle(contours[j])
                 centerj = (int(xj),int(yj))
                 radiusj = int(radiusj)
                 #checking here
                 if (centerj[0] == center[0]) | (centerj[0] in range(center[0]-5 , center[0]+5)):
                    temp.append(centerj)
       
           #temp.reverse()
           ans[AvAns[k]]+=temp
           AnsRef["col "+str(i)]=ans
        
    r=radiusj
    print "r= ", r  #r here is integer
        
    contours1, hierarchy = cv2.findContours (sheetTh["SecBn"],cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
    contours1.reverse() 
    #cv2.imshow("coc",sheetTh["SecBn"])     
    for n in range(10):
       (x,y),r=cv2.minEnclosingCircle(contours1[n])
       x=int(x)
       SBnRef[n]=x
   
    print "AnsRef= ",AnsRef
    print "SBnRef= ",SBnRef
    return AnsRef,SBnRef
''' -----------------------------------------------------------------------'''

''' now get the correct answers and extract their positions '''
''' user enter correct ansers, then extract postions, and finally 
create a threshold model for them '''

def Cor_Ans(NQ,AvAns,AnsRef,Q,QpC,CAL):
    
    print "rCall=" ,r
    ri=int(r)  #so woerd issue here, r is float while it's integer in the previous step
    
    for q in range(NQ):
        Ques_Stat["Q"+str(q+1)]={}
        Ans_Stat["Q"+str(q+1)]={}
        for c in case:
            Ques_Stat["Q"+str(q+1)][c]=0
        for a in AvAns:
            Ans_Stat["Q"+str(q+1)][a]=0
    #Ans_stat was build but not used.       
    print "QuesStat_B= ", Ques_Stat
    print "AnsStat_B= ", Ans_Stat
    
    print "CAL_call= ", CAL
    CorAns["col 0"]=CAL[0:QpC]
    CorAns["col 1"]=CAL[QpC:2*QpC]
    CorAns["col 2"]=CAL[2*QpC:3*QpC]
    CorAns["col 3"]=CAL[3*QpC:]
    
    for n in range(4):
        Ps=[]
        for i in range(Q[n]):  #no. of Q in each col
            lenAns = len(CorAns[cols[n]][i])
            if lenAns > 1 :
                tem=[]
                tem.append(AnsRef[cols[n]][CorAns[cols[n]][i][0]][i])
                tem.append(AnsRef[cols[n]][CorAns[cols[n]][i][1]][i])
                Ps.append(tem)

            else:
                Ps.append(AnsRef[cols[n]][CorAns[cols[n]][i][0]][i])

        CorAnsPos[cols[n]]=Ps
  
     
    for c in cols:
        for an in CorAnsPos[c]:
            lan=str(an)
            lan=lan.count(",")-1
            if lan > 1:
                for ct in range(lan):
                    cv2.circle(sheet[c],an[ct],ri,(100,15,120),-1)
            else:
                cv2.circle(sheet[c],an,ri,(100,15,120),-1)
        #cv2.imshow("CorrectAns_Normal-"+c,sheet[c])
        cv2.imwrite("CorrectAns_Normal-"+c+".png",sheet[c])
      
        CaGry=cv2.cvtColor(sheet[c],cv2.COLOR_BGR2GRAY)
        CaR, Cath = cv2.threshold(CaGry, 100, 255, cv2.THRESH_BINARY_INV)
        #cv2.imshow("CorrectAns_TH-"+c,Cath)
        cv2.imwrite("CorrectAns_TH-"+c+".png",Cath)
        
        Cath=cv2.medianBlur(Cath,7)
        #cv2.imshow("CorrectAns_TH/B-"+c,Cath)
        cv2.imwrite("CorrectAns_TH_B-"+c+".png",Cath)
           
        CAth[c]=Cath
        
        #cv2.imshow(c,Cath)
        
    print "QuesStat= ", Ques_Stat
    print "AnsStat= ", Ans_Stat
    print "CorAns= ", CorAns
    print "CorAnsPos= ", CorAnsPos
    print "CAth= ", CAth
    return Ques_Stat, Ans_Stat, CorAns, CorAnsPos, CAth 
 
''' ---------------------------------------------------------------------'''


def ROIStu(imgstu,cen):
    StuSheet={}
    StuSheet2={}
    gray= cv2.cvtColor(imgstu,cv2.COLOR_BGR2GRAY)
    th, im_th = cv2.threshold(gray, 140, 255, cv2.THRESH_BINARY_INV)
    #cv2.imshow("student",im_th)
    cv2.imwrite("Student_TH.png",im_th)
    im_out = cv2.medianBlur(im_th,11)
    #cv2.imshow("student",im_out)
    cv2.imwrite("Student_TH_B.png",im_out)
    im2=im_out.copy()
    
    for i in cen:  
       y1=int(cen[i][0][1])
       y2=int(cen[i][1][1])
       x1=int(cen[i][0][0])
       x2=int(cen[i][2][0])
       col=im_out[y1:y2,x1:x2]       
       StuSheet[i]=col
       col2=im2[y1:y2,x1:x2]
       StuSheet2[i]=col2
       #cv2.imshow("Stu-"+i, StuSheet[i])
       cv2.imwrite("Student-"+i+".png", StuSheet[i])
    return StuSheet, StuSheet2
''' -------------------------------------------------------------------'''

def StuAnsPos(StuSheet):
    for ck in cols:
        stucon, hi = cv2.findContours(StuSheet[ck],cv2.RETR_TREE,cv2.CHAIN_APPROX_SIMPLE)

        sans={}
        temp=[]
        for i in range(len(stucon)):
            (x,y),radius = cv2.minEnclosingCircle(stucon[i])
            center = (int(x),int(y))
            radius = int(radius)
            temp.append(center)
        temp.reverse()    
        sans[ck]=temp
    
        lsans=len(sans[ck])-1
        for aa in range(lsans):
          try:
            t=[]
            if (sans[ck][aa][1] == sans[ck][aa+1][1]) | (sans[ck][aa][1] in range(sans[ck][aa+1][1]-5 , sans[ck][aa+1][1]+5)) :
                t.append(sans[ck][aa])
                t.append(sans[ck][aa+1])
                sans[ck][aa]=t
                sans[ck].remove(sans[ck][aa+1])
  
          except IndexError:
             break
    print "St_Ans= ", sans
    return sans
''' ------------------------------------------------------------------'''
         
def Grading(StuSheet,CAth):
    res={}
    res2={}
    for c in cols:
        res[c]= StuSheet[c] & CAth[c]
        #cv2.imshow("res 0",res[c])
        cv2.imwrite("And_Res-"+c+".png",res[c])
        
        res2[c]=res[c].copy()
        res_con, res_hier=cv2.findContours(res[c],cv2.RETR_TREE,cv2.CHAIN_APPROX_SIMPLE)
        temp=[]
        for con in res_con:
            (X,Y), Rad=cv2.minEnclosingCircle(con)
            CA_cen=(int(X),int(Y))
            temp.append(CA_cen)
            
        temp.reverse()
        Stu_Cor_Ans[c]=temp


        ''' in case two answers'''

        lt=len(Stu_Cor_Ans[c])-1
        for aa in range(lt):
            try:
                tl=[] #temp list
                if (Stu_Cor_Ans[c][aa][1] == Stu_Cor_Ans[c][aa+1][1]) | (Stu_Cor_Ans[c][aa][1] in range(Stu_Cor_Ans[c][aa+1][1]-10 , Stu_Cor_Ans[c][aa+1][1]+10)) :
                    tl.append(Stu_Cor_Ans[c][aa])
                    tl.append(Stu_Cor_Ans[c][aa+1])
                    
                    Stu_Cor_Ans[c][aa]=tl
                    Stu_Cor_Ans[c].remove(Stu_Cor_Ans[c][aa+1])
            except IndexError:
                break
   
    print "Stu_Cor_Ans= ",Stu_Cor_Ans        
    return Stu_Cor_Ans,res2
    
''' ------------------------------------------------------------------'''

def CorWroMisQ(res2,sheetTh2,sheetTh3,StuSheet2,NAv,CAth,NQ,CorAns,CAG,neg_G,ONT):
    
    
    WQth={}
    Stu_Qs_State={}
    def getqcase(y,x,NAv,nts):  #nts: name to save the image with
       qul2=[]
       rest={}
       qcl={}
       qulc={}
       for c in cols:
           #print y.keys()
           #cv2.imshow("y",y["col 3"])
           rest[c]= (y[c] & cv2.bitwise_not(x[c])) | (cv2.bitwise_not(y[c]) & x[c])   
           cv2.imwrite(nts+"_TH"+c+".png",rest[c])
           rest[c]=cv2.medianBlur(rest[c],11) 
           cv2.imwrite(nts+"_TH_B-"+c+".png",rest[c])
           
           res_con2, res_hier=cv2.findContours(rest[c],cv2.RETR_TREE,cv2.CHAIN_APPROX_SIMPLE)
           
           qul=[]
           t3=[]
           for s in res_con2:
               (xt,yt),radt = cv2.minEnclosingCircle(s)
               centt = (int(xt),int(yt))
               qul.append(centt)
         
           qul.reverse() 
           qcl[c]=qul
           l=len(qcl[c])-1
           rc=0
           for m in range(l):
               if (qcl[c][m][1] == qcl[c][m+1][1]) | (qcl[c][m][1] in range (qcl[c][m+1][1]-2, qcl[c][m+1][1]+3)):
                   rc+=1
                   
               elif rc==NAv-1:
                   
                   for r in range(NAv):
                        t3.append(qcl[c][m-r])
                   #print "list  ", t3
                   qul2.append(t3)
                   t3=[]
                   rc=0
                   
               else:
                   
                   for r in range(rc+1):
                       t3.append(qcl[c][m-r])  
                   
                   qul2.append(t3)
                   #print "list  ", t3
                   '''
                   print "--------- " +c
                   print t3
                   print "------------" 
                   '''
                   t3=[]
                   rc=0
           
           for r in range(rc+1):
                       t3.append(qcl[c][m-r+1])  
                   
           qul2.append(t3)
           #print "list  ", t3
           qulc[c]=qul2 
           qul2=[]
       
       #print len(qul2)
       return qulc        

    for j in cols:
         Stu_Qs_State[j]=[]
    print "Stu_Qs_State_main= ", Stu_Qs_State

    
    MisQ=getqcase(StuSheet2,sheetTh3,NAv,"XOR_Mis") #string to savee the resault only.
    CorQ=getqcase(res2,sheetTh2,NAv,"XOR_Cor")   # string to save the results only
    
    print "MisQ= ", MisQ
    print "CorQ= " ,CorQ
    
    grade=0
    for c in cols :
        for h in CorQ[c]:
            if len(h) < NAv :
                Stu_Qs_State[c].append(NAv-len(h))
            else:
                Stu_Qs_State[c].append("--")
        print "Stu_Qs_State_1st_step= ", Stu_Qs_State
    
        for t in range(len(Stu_Qs_State[c])):
            LSCA=len(CorAns[c][t])
            LSSA=Stu_Qs_State[c][t]  #student answer as astring
      
            if (LSSA == LSCA) & (Stu_Qs_State[c][t] != "--"):
                Stu_Qs_State[c][t]= "T"
                
            # here one instead of two 
            elif Stu_Qs_State[c][t] != "--":
                Stu_Qs_State[c][t]= "P"
            
            print "Stu_Qs_State_2nd_step= ", Stu_Qs_State
            
            for M in MisQ[c] :
                if len(M) == NAv :
                    Stu_Qs_State[c][MisQ[c].index(M)]="M"
                    
            print "Stu_Qs_State_3rd_step= ", Stu_Qs_State


        for w in Stu_Qs_State[c] :
            if w == "--" :
                Stu_Qs_State[c][Stu_Qs_State[c].index(w)]="W" 
                
        print "Stu_Qs_State_4th_step= ", Stu_Qs_State
        
        ''' correct and wrong mark'''
        for q in Stu_Qs_State[c]:
            if q == "T":
                grade+=CAG
            elif q == "W":
                grade-=neg_G
            elif q == "P":
                grade+=ONT
            else:
                None
    '''        
    print CorQ
    print "----------------------------------"
    print WQ
    print "=========================================="
    print MisQ
    print "-------------------------------------"
    '''
    print "Stu_Qs_State_final= ", Stu_Qs_State
    return Stu_Qs_State,grade 

            
''' ----------------------------------------------------------------------'''

def GetSecBn(SBnRef, StuSheet2,grade,Results):
    stid=StuSheet2["SecBn"]
    
    contours2, hierarchy = cv2.findContours (stid,cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
    contours2.reverse()

    sid=[]
    ID=[]
    for c in range(len(contours2)):
        (x2,y),r=cv2.minEnclosingCircle(contours2[c])
        x2=int(x2)
        sid.append(x2)
    print "sid= " ,sid
    
    for s in sid:
        for n in SBnRef:
            if (s == SBnRef[n]) | (s in range(SBnRef[n]-5, SBnRef[n]+5)):
                ID.append(n)
                break

    ID=(ID[0],10*ID[1]+ID[2])
    print "ID= ", ID
    Results[str(ID[0])][ID[1]-1]=grade
    print "Results= " ,Results
    
    return ID,Results
'''-------------------------------------------------------------------------'''

def Qus_Statistics(Stu_Qs_State,Ques_Stat):
    
    Q=0
    for c in colsr: 
        for q in Stu_Qs_State[c]:
            Q+=1
            Qk= "Q"+ str(Q) # question key to inter Ques_Stat
            
            
            if q == "T" or q== "P":
                Ques_Stat[Qk]["C"]+=1
            
            elif q == "W":
                Ques_Stat[Qk]["W"]+=1
            
            elif q == "M":
                Ques_Stat[Qk]["M"]+=1
            
    print "Ques_Stat= ", Ques_Stat
    return Ques_Stat
    
''' -----------------------------------------------------------'''
def Final_Report(Ques_Stat,Results,title,NQ,nSec,stpsec,CAG):
    
    book = xlwt.Workbook(encoding="utf-8")
    sh1 = book.add_sheet("Grades Sheet")
    sh2 = book.add_sheet("Ques Statistics")
    
    sh1.write(0, 0, 'Sec')
    sh1.write(0, 1, 'BN')
    sh1.write(0, 2, 'Mark ('+ str(NQ*CAG)+ ')' )
     
    y=stpsec
    
    for s in range(nSec):
        
        for bn in range(stpsec):
            r=s+1
            sh1.write(bn+1+y*(r-1), 0, s+1)
            n=bn+1
            sh1.write(bn+1+y*(r-1), 1, n)
            
            g=Results[str(r)][bn]
            
            sh1.write(bn+1+y*(r-1), 2, g)  #bn here represente grade
    

    n_ques = NQ
    qn=[]
    cor=[]
    wor=[]
    mis=[]

    for q in range(n_ques):
        t="Q"+ str(q+1)
    
        qn.append(t)
    
        cor.append(Ques_Stat[t]["C"])
        wor.append(Ques_Stat[t]["W"])
        mis.append(Ques_Stat[t]["M"])
        
        sh2.write(q+1,0,t)
        
    
    
    sh2.write(0,0,'Question')
    sh2.write(0, 1, 'Correct')
    sh2.write(0, 2, 'Wrong')
    sh2.write(0, 3, 'Missing')
    
    for q in range(NQ):
        sh2.write(q+1,1,cor[q])
        
        sh2.write(q+1,2,wor[q])
        
        sh2.write(q+1,3,mis[q])   
        
    book.save(title+".xls")
    
   