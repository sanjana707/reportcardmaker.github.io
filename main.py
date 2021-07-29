#!/usr/bin/env python
# coding: utf-8


import sys
#get_ipython().system('{sys.executable} -m pip install openpyxl')


#get_ipython().system('{sys.executable} -m pip install reportlab')


import openpyxl


import glob


from reportlab import *
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.pdfmetrics import stringWidth

from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.colors import HexColor

import pandas as pd

#global variables
width, height = A4   #595.2755905511812 841.8897637795277


def get_data():
    file = 'Dummy Data.xlsx'
    data = pd.ExcelFile(file).parse()
    
    ps = openpyxl.load_workbook(file, data_only=True)
    sheet = ps['Sheet1']
    
    details=dict()
    
    
    for row in range(3, sheet.max_row+1):
        reg_no=sheet['F'+str(row)].value
        #print(f_name)
        
        if reg_no not in details.keys():
            details.update({reg_no: {}})
            
            #details[reg_no]=reg_no
            details[reg_no]["rnd"]=sheet['B'+str(row)].value
            details[reg_no]["f_name"]=sheet['C'+str(row)].value
            details[reg_no]["l_name"]=sheet['D'+str(row)].value
            details[reg_no]["full_name"]=sheet['E'+str(row)].value
            details[reg_no]["grade"]=sheet['G'+str(row)].value
            details[reg_no]["school"]=sheet['H'+str(row)].value
            details[reg_no]["gender"]=sheet['I'+str(row)].value
            details[reg_no]["dob"]=sheet['J'+str(row)].value
            details[reg_no]["city"]=sheet['K'+str(row)].value
            details[reg_no]["date"]=sheet['L'+str(row)].value
            details[reg_no]["country"]=sheet['M'+str(row)].value
            details[reg_no]["final"]=sheet['T'+str(row)].value
            
            q_no=[sheet['N'+str(row)].value]
            ans_marked=[sheet['O'+str(row)].value]
            ans_correct=[sheet['P'+str(row)].value]
            outcome=[sheet['Q'+str(row)].value]
            max_score=[sheet['R'+str(row)].value]
            score=[sheet['S'+str(row)].value]
            
            details[reg_no]["q_no"]=q_no
            details[reg_no]["ans_marked"]=ans_marked
            details[reg_no]["ans_correct"]=ans_correct
            details[reg_no]["outcome"]=outcome
            details[reg_no]["max_score"]=max_score
            details[reg_no]["score"]=score
            
        else:
            details[reg_no]["q_no"].append(sheet['N'+str(row)].value)
            details[reg_no]["ans_marked"].append(sheet['O'+str(row)].value)
            details[reg_no]["ans_correct"].append(sheet['P'+str(row)].value)
            details[reg_no]["outcome"].append(sheet['Q'+str(row)].value)
            details[reg_no]["max_score"].append(sheet['R'+str(row)].value)
            details[reg_no]["score"].append(sheet['S'+str(row)].value)
            
    return details


def plot_data(details):
    for key in details.keys():
        #print(key)
        s = Student(details[key], key)
        s.create_pdf()


class Student:
    def __init__(self, reg_no, actual_reg_no):
        self.reg_no=actual_reg_no
        self.dict=reg_no   #dictionary of details
        self.f_name = reg_no['f_name']
        self.l_name = reg_no['l_name']
        self.full_name = reg_no['full_name']
        self.rnd = reg_no['rnd']
        self.grade = reg_no['grade']
        self.school = reg_no['school']
        self.gender = reg_no['gender']
        self.dob = reg_no['dob']
        self.date = reg_no['date']
        self.country = reg_no['country']
        self.final = reg_no['final']
        
        self.q_no=reg_no['q_no']
        self.ans_marked=reg_no['ans_marked']
        self.ans_correct=reg_no['ans_correct']
        self.outcome=reg_no['outcome']
        self.max_score=reg_no['max_score']
        self.score=reg_no['score']
        
        #student image
        self.image='Pics_for_assignment/'+reg_no['full_name']+'.png'
        
    def get_attributes(self):
        print(vars(self))
        
    
    def insert_logo(self,can):
        #for logo
        logo_file="school_logo.jpg"
        logo_width=100
        x=40
        y=height-100
        can.drawImage(logo_file, 40, height-100, width=120, height=120, mask='auto')
         
    def insert_school(self,can):
        #for School name
        can.setFont('Helvetica', 50)
        can.setFillColor((HexColor('#144c5e')))
        text_width=stringWidth(self.school, 'Helvetica', 50)
        x=200
        y=height-50  
        can.drawString(x, y, self.school)
        can.line(x,y-2, x+text_width,y-2)
        
    def insert_line(self, can, x, y):
        can.line(0,y, width, y)
        
    def insert_student(self, can, simg_width, simg_height):    
        x=width-simg_width-30
        y=height-5-simg_height
        can.drawImage(self.image, x,y, width=simg_width,height=simg_height, preserveAspectRatio=True, mask='auto')
        
        #caption
        can.setFillColor((HexColor('#071b21')))
        can.setFont('Helvetica', 10)
        can.drawString(x+20,y-10,self.full_name)
        
    def header(self, can, simg_width, simg_height):
        #report card tag-line
        text_width=stringWidth("Report Card", 'Helvetica', 20)
        x=(width-text_width)/2
        y=height-110
        can.setFillColor((HexColor('#071b21')))
        can.setFont('Helvetica-Bold', 20)
        can.drawString(x,y, "Report Card")
        
        #round
        text_width=stringWidth("Round: "+str(self.rnd), 'Helvetica', 12)
        x=(width-text_width)/2
        y=height-130
        can.setFont('Helvetica', 12)
        can.drawString(x,y, "Round: "+str(self.rnd))
        
    def create_data_table(self, can): 
        #create table
        
        data=[["Registration No. ", self.reg_no,"Grade", self.grade],
              ["First Name", self.f_name, "Date of Birth", self.dob],
              ["Last Name", self.l_name, "Counrty", self.country],
              ["Gender", self.gender, "Date & Time of Test", self.date]]
        
        tab_width = 400   #500
        tab_height = 200  #200
        x = (width-tab_width)/2
        y = height-220
        
        t = Table(data)
        s=[('GRID',(0,0),(-1,-1),0.5,colors.gray)]
        can.setFillColor((HexColor('#071b21')))
        t.setStyle(TableStyle(s))
        t.wrapOn(can, tab_width, tab_height)
        t.drawOn(can, x, y)
        
    def show_result(self, can):
        total=0
        max_score=0
        for i in self.dict["score"]:
            total+=int(i)
        for j in self.dict["max_score"]:
            max_score+=int(j)
        
        x=100
        y=height-250
        
        can.setFillColor((HexColor('#144c5e')))
        can.setFont('Helvetica-Bold', 15)
        
        text_width=stringWidth("Your Score: "+str(total)+"/"+str(max_score), 'Helvetica-Bold', 15)
        x=(width-text_width)/2
        can.drawString(x,y,"Your Score: "+str(total)+"/"+str(max_score))
        
        text_width=stringWidth(self.final, 'Helvetica-Bold', 15)
        x=(width-text_width)/2
        can.drawString(x, y-30, self.final)
        
        
    def create_marks_table(self, can):
        data=[[]]
        #print(data)
        
        data[0]=["Question No.", "Answer Marked", "Correct Answer", 
                 "Outcome","Score if Correct", "Your Score"]
        
        
        for i in range(25):
            data.append([self.dict["q_no"][i], 
                       self.dict["ans_marked"][i],
                       self.dict["ans_correct"][i],
                       self.dict["outcome"][i],
                       self.dict["max_score"][i],
                       self.dict["score"][i]])
        
        tab_width = 500
        tab_height = 400
        x = 80
        y = height-780
        
        s=[('GRID',(0,0),(-1,-1),0.5,colors.gray),
             ('ALIGN',(0,1),(-1,-1),'CENTER'),
              ('BACKGROUND',(0,0),(5,0), colors.lightblue )]
        can.setFillColor((HexColor('#071b21')))
        t = Table(data)
        t.setStyle(TableStyle(s))
        t.wrapOn(can, tab_width, tab_height)
        t.drawOn(can, x, y)
              
    def create_pdf(self):
        file = self.full_name+'.pdf'
        can = canvas.Canvas(file, pagesize=A4)
        
        #width, height = A4   #595.2755905511812 841.8897637795277
        
        self.insert_logo(can)
        self.insert_school(can)
        #self.insert_line(can, 0, height-80)
        
        simg_width=100
        simg_height=120
        self.insert_student(can, simg_width, simg_height)
        
        self.header(can, simg_width, simg_height)
        self.create_data_table(can)
        
        self.show_result(can)
        
        self.create_marks_table(can)
        
        can.showPage()
        can.save()
        
        #print(self.image)
        
        


if __name__=="__main__":
    details=get_data()
    #print(details)
    
    plot_data(details)
    