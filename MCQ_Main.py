# -*- coding: utf-8 -*-
"""
This Application is to create answer sheets for MCQ exams, 
and also to mark those exams.
"""

from kivy.app import App
from kivy.lang import Builder
from kivy.uix.screenmanager import ScreenManager, Screen, SlideTransition
from kivy.uix.widget import Widget
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox
from kivy.uix.togglebutton import ToggleButton
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.image import Image
from kivy.uix.popup import Popup
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.scrollview import ScrollView
from filebrowser import FileBrowser
from kivy.properties import ObjectProperty
from os.path import sep, expanduser, isdir, dirname
from kivy.core.window import Window
import os
import glob
import string
import MCQ_Lib
import cv2

Window.size = (650, 550) #set the window size


''' First screen, user here select task'''
class WelcomeScreen(Screen):
    pass


'''
New sheet setup screen
w,x,y,z represents user's inputs;
   w: no. of questios,
   x: no. of availble choices,
   y: title,
   z: class
these inputs should be checked first to make sure nothing 
is wrong or missing
'''
class CreateSheet(Screen):
    def check(self,w,x,y,z):
        
        if w == "" or w.isspace()==True:
            self.ids.nq.hint_text_color=[10,0,0,1]
            self.ids.nq.text=""
            self.ids.nq.hint_text= "**Required"
            self.ids.nq.background_color=[0.9,0.9,0.9,1]
    
        elif int(w) > 160:
            self.ids.nq.hint_text_color=[10,0,0,1]
            self.ids.nq.text=""
            self.ids.nq.hint_text= "Max no. is 160"
            self.ids.nq.background_color=[0.9,0.9,0.9,1]
        
          
        if x == "" or x.isspace()==True:
            self.ids.nc.hint_text_color=[10,0,0,1]
            self.ids.nc.text=""
            self.ids.nc.hint_text= "**Required"
            self.ids.nc.background_color=[0.9,0.9,0.9,1]   
            
        elif int(x) > 5:
            #print type(x)
            self.ids.nc.hint_text_color=[10,0,0,1]
            self.ids.nc.text=""
            self.ids.nc.hint_text= "Max no. is 5"
            self.ids.nc.background_color=[0.9,0.9,0.9,1]
            
            
        if y == "" or y.isspace()==True:
            self.ids.ttl.hint_text_color=[10,0,0,1]
            self.ids.ttl.text=""
            self.ids.ttl.hint_text= "**Required"
            self.ids.ttl.background_color=[0.9,0.9,0.9,1]
            
            
        if z == "" or z.isspace()==True:
            self.ids.clss.hint_text_color=[10,0,0,1]
            self.ids.clss.text=""
            self.ids.clss.hint_text= "**Required"
            self.ids.clss.background_color=[0.9,0.9,0.9,1]
        
        else:        
            MCQ_Lib.Ans_Sheet(self,w,x,y,z)
'''---------------------------------------------------------------------'''

'''for loading files'''
class LoadDialog(FloatLayout):
    load = ObjectProperty(None)
    cancel = ObjectProperty(None)


'''
here, user set grading method, upload empty answersheet,
and a file contains students'sheets and all as PNG files.
**NOTE**
all images should be pre-processed, i.e. no tilting and no spots.
'''
class GradingMethod(Screen):
    f=ObjectProperty(None)
    def dismiss_popup(self):
        self._popup.dismiss()

    def show_load(self):
        content = LoadDialog(load=self.load, cancel=self.dismiss_popup)
        self._popup = Popup(title="Load file", content=content,
                            size_hint=(0.9,0.9),auto_dismiss=False)
        self._popup.open()

    
    def load(self,filename):
            self.f.text = str(filename[0])
            #print filename
            self.dismiss_popup()
    '''*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-**-*-*-*-*-*-*-*-*-*-*'''     
    
    '''
    here, inputs should be checked first to make sure nothing 
    is wrong or missing.
    info is a list containing ids for all inputs.
    **check at the KV file.
    ''' 
    def check(self,info):
        i=0  #counter to check input data
        
        '''//check no. of questions'''
        try:
            if(int(self.ids.nq.text) > 160):
                self.ids.nq.hint_text_color=[10,0,0,1]
                self.ids.nq.text=""
                self.ids.nq.hint_text= "Max no. is 160"
                self.ids.nq.background_color=[0.9,0.9,0.9,1]
                
            else:
                i+=1
           
        except ValueError:
            self.ids.nq.text=""
            self.ids.nq.hint_text_color=[10,0,0,1]
            self.ids.nq.hint_text= "**Required"
            self.ids.nq.background_color=[0.9,0.9,0.9,1]
                
 
        '''//check no. of choices'''
        try:
            if int(self.ids.nc.text) > 5:
                self.ids.nc.hint_text_color=[10,0,0,1]
                self.ids.nc.text=""
                self.ids.nc.hint_text= "Max no. is 5"
                self.ids.nc.background_color=[0.9,0.9,0.9,1]
                
            else:
                i+=1
        
        except ValueError:
            self.ids.nc.text=""
            self.ids.nc.hint_text_color=[10,0,0,1]
            self.ids.nc.hint_text= "**Required"
            self.ids.nc.background_color=[0.9,0.9,0.9,1]
        
        
        '''//check no. of mark for correct answer'''
        try:
            int(self.ids.cag.text)
            i+=1
            
        except ValueError:
            self.ids.cag.text=""
            self.ids.cag.hint_text_color=[10,0,0,1]
            self.ids.cag.hint_text= "**Required"
            self.ids.cag.background_color=[0.9,0.9,0.9,1]
            i-=1
     
        '''//check negative grading'''
        if self.ids.cwn.state == 'down':
            i+=1
            '''//check -ve mark', it should be +ve value'''
            
            try:
                if float(self.ids.ng.text) < 0:
                    self.ids.ng.hint_text_color=[10,0,0,1]
                    self.ids.ng.text=""
                    self.ids.ng.hint_text= "Postive no. only"
                    self.ids.ng.background_color=[0.9,0.9,0.9,1]
                    i-=1
                     
            except ValueError:
                self.ids.ng.text=""
                self.ids.ng.hint_text_color=[10,0,0,1]
                self.ids.ng.hint_text= "**Required"
                self.ids.ng.background_color=[0.9,0.9,0.9,1]
                i-=1
        
        else:
            i+=1
            
            
        '''//check mark for one answer instead of two'''
        try:
            float(self.ids.ont.text)
            i+=1
            
        except ValueError:
            self.ids.ont.text=""
            self.ids.ont.hint_text_color=[10,0,0,1]
            self.ids.ont.hint_text= "**Required"
            self.ids.ont.background_color=[0.9,0.9,0.9,1]
            i-=1
        
        '''//check no. of sections, max no. is nine.'''
        try:
            if int(self.ids.nsec.text) > 9:
                self.ids.nsec.hint_text_color=[10,0,0,1]
                self.ids.nsec.text=""
                self.ids.nsec.hint_text= "Max no. is 9"
                self.ids.nsec.background_color=[0.9,0.9,0.9,1]
                
            else:
                i+=1
                
        except ValueError:
            self.ids.nsec.text=""
            self.ids.nsec.hint_text_color=[10,0,0,1]
            self.ids.nsec.hint_text= "**Required"
            self.ids.nsec.background_color=[0.9,0.9,0.9,1]
            
            
        '''//check no. of student in each section, max no. is 99.'''    
        try:
            if int(self.ids.stu.text) > 99:
                self.ids.stu.hint_text_color=[10,0,0,1]
                self.ids.stu.text=""
                self.ids.stu.hint_text= "Max no. is 99"
                self.ids.stu.background_color=[0.9,0.9,0.9,1]
                
            else:
                i+=1
                
        except ValueError:
            self.ids.stu.text=""
            self.ids.stu.hint_text_color=[10,0,0,1]
            self.ids.stu.hint_text= "**Required"
            self.ids.stu.background_color=[0.9,0.9,0.9,1]
            
        
        '''//check title to save the final result.'''
        t=self.ids.ttls.text   
        if t.isspace() == True or t=="":
            self.ids.ttls.text=""
            self.ids.ttls.hint_text_color=[10,0,0,1]
            self.ids.ttls.hint_text= "**Required"
            self.ids.ttls.background_color=[0.9,0.9,0.9,1]
        
        else:
            i+=1
        
        
        '''//check that empty sheet is uploaded'''
        es=self.ids.emt.text
        if es.isspace() == True or es=="":
            self.ids.emt.text=""
            self.ids.emt.hint_text_color=[10,0,0,1]
            self.ids.emt.hint_text= "**Required"
            self.ids.emt.background_color=[0.9,0.9,0.9,1]
            
        else:
            i+=1
         
         
        '''//check that students folder is uploaded''' 
        sf=self.ids.sf.text
        if sf.isspace() == True or sf=="":
            self.ids.sf.text=""
            self.ids.sf.hint_text_color=[10,0,0,1]
            self.ids.sf.hint_text= "**Required"
            self.ids.sf.background_color=[0.9,0.9,0.9,1]
            
        else:
            i+=1
            #print i
        '''*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'''
        
        ''' here, four functions are excuted to build base refrences to 
            be used later in marking operation'''
            
        if i == 10:
            global NQ,NAv,CAG,minus_G,neg_G,ONT,nSec,stpsec,title,EMT,SF,AvAns,Q,QpC,emptysheet,stsheets,Results
        
            NQ,NAv,CAG,minus_G,neg_G,ONT,nSec,stpsec,title,EMT,SF,AvAns,Q,QpC,emptysheet,stsheets=MCQ_Lib.SheetInfo(self,info)
        
            Results=MCQ_Lib.RDic(nSec,stpsec) 
            
            global sheet, sheetTh, sheetTh2,sheetTh3, cen, AnsRef,SBnRef
            
            cen, sheet, sheetTh, sheetTh2,sheetTh3=MCQ_Lib.EmptySheet(emptysheet)
            AnsRef,SBnRef=MCQ_Lib.Extract_Pos(sheetTh,NAv,AvAns)
            
            self.manager.transition = SlideTransition(direction="left")
            self.manager.current = "cor ans"
            
        
            #print NQ,NAv,CAG,minus_G,neg_G,ONT,nSec,stpsec,title,EMT,SF,AvAns,Q,QpC,Results


'''
user here enters correct answers, mark and get final results'''
class AnsMarkScreen(Screen):
    
    '''flags to make sure every function had been excuted correctly'''
    Lf=False #loading answers board
    Sf=False #Save answers
    Mf=False #Mark sheets
    Rf=False #Results
    
    ''' build board to enter answers'''
    def buildans(self):      
        if self.Lf == False:
            global dic
            dic= {}

            
            box=BoxLayout(size_hint=(self.ids.bx21.width,0.1))
            self.ids.grd.add_widget(box)
           
            for i in range(NQ):
                dic[str(i+1)]=[]
                box=BoxLayout()
                q=Label(size_hint_y=0, height=20,text=str(i+1))
                box.add_widget(q)
                
                for c in AvAns:             
                    cbx = ToggleButton(size_hint_y=0, height=20,text=c) 
                    dic[str(i+1)].append(cbx)
                    box.add_widget(cbx)
                
                self.ids.grd.add_widget(box)
                
            self.Lf=True
            
    '''check toggelbutton states to know answers'''
    def states(self,*args):
        global CAL
        CAL=[]
        self.Sf=False
        if self.Lf==True:
            for j in range(NQ):
                CAL.append([])
                
            '''store correct answers in a dictionary'''
            for k in dic:
                for a in dic[k]:
                    if a.state=="down":
                        CAL[int(k)-1].append(AvAns[dic[k].index(a)])
            #print "CAL= ",CAL
                        
        else:
            con= BoxLayout()
            L=Label(text="Please enter the answers first")
            con.add_widget(L)
            self._popup = Popup(title="Warning", content=con,
                        size_hint=(0.7,0.2),auto_dismiss=True)
            self._popup.open()
        
        print "CAL= ", CAL
    '''getting correct answers' positions, and build a dictionary 
       to know questions'states later'''
    def CorrAnss(self,*args):
        if self.Lf==True:
            try:
                global Ques_Stat, Ans_Stat, CorAns, CorAnsPos, CAth
                Ques_Stat, Ans_Stat, CorAns, CorAnsPos, CAth=MCQ_Lib.Cor_Ans(NQ,AvAns,AnsRef,Q,QpC,CAL)
                self.Sf=True
                
            except IndexError:
                
                con= BoxLayout()
                L=Label(text="Make sure that all questions have been answered")
                con.add_widget(L)
                self._popup = Popup(title="Warning", content=con,
                            size_hint=(0.7,0.2),auto_dismiss=True)
                self._popup.open()
                
                
                #print "ncal", CAL
        #print Ques_Stat
        

    ''' marking starts here;
        get each sheet from the folder, split it into five rejions,
        then start to compare answers, 
        calculate final grade,
        get the ID,
        update Reslts dictionary,
        and finally update statistics dictionary.
        
    '''
    def Student(self,*args):
        
        if self.Sf ==True:
            global Ques_Stat
            
            path = stsheets
    
            for f in glob.glob(os.path.join(path, '*.png')):
            
                imgstu=cv2.imread(f)
            
           
                StuSheet,StuSheet2=MCQ_Lib.ROIStu(imgstu,cen)
            
                Stu_Cor_Ans,res2=MCQ_Lib.Grading(StuSheet,CAth)
                sans=MCQ_Lib.StuAnsPos(StuSheet)
            
                Stu_Qs_State,grade=MCQ_Lib.CorWroMisQ(res2,sheetTh2,sheetTh3,StuSheet2,NAv,CAth,NQ,CorAns,CAG,neg_G,ONT)
                
                #print "NQ=" ,NQ
                #print "Results=", Results
                ID,Result=MCQ_Lib.GetSecBn(SBnRef, StuSheet2,grade,Results)
                Ques_Stat=MCQ_Lib.Qus_Statistics(Stu_Qs_State,Ques_Stat)
           
                #print "ID= ",ID, " grade= " ,grade
                #print Stu_Qs_State
                
                StuSheet,StuSheet2,Stu_Cor_Ans,res2,sans,grade,WQ,CorQ,MisQ,Stu_Qs_State,ID= [0,0,0,0,0,0,0,0,0,0,0]
                
                self.Mf= True
       
        else:
            con= BoxLayout()
            L=Label(text="Please Save the answers first")
            con.add_widget(L)
            self._popup = Popup(title="Warning", content=con,
                        size_hint=(0.7,0.2),auto_dismiss=True)
            self._popup.open()
            
        
    ''' creating the final file that contains results and statistics'''        
    def resultfile(self,*args):
        if self.Mf==True:
            Gfile=MCQ_Lib.Final_Report(Ques_Stat,Results,title,NQ,nSec,stpsec,CAG)
            con= BoxLayout()
            L=Label(text="Process has been finshed")
            con.add_widget(L)
            self._popup = Popup(title="Final Result", content=con,
                                size_hint=(0.7,0.2),auto_dismiss=False)
            self._popup.open()
            
        else:
            con= BoxLayout()
            L=Label(text="Please enter the answers, save them and press mark first")
            con.add_widget(L)
            self._popup = Popup(title="Warning", content=con,
                        size_hint=(0.7,0.2),auto_dismiss=True)
            self._popup.open()
            
            
    
        #print " Final"
        #print Ques_Stat
        #print Results
'''*_*_*_*_*_*_*_*_*_*_*_*_*__*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*__*_*_*_*'''        
      
class ScreenManagement(ScreenManager):
    pass


buildUI = Builder.load_file("MCQ_UI.kv")

class MCQ_UI(App):
    def build(self):
        return buildUI

MCQ_UI().run()
