#=====================================Imports=============================

import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import *
#from PyQt5 import QtCore, QtWidgets
import openpyxl
from openpyxl import *
from openpyxl import Workbook
from xlsxwriter import Workbook
from tkinter import messagebox
from tkinter import SCROLL
from tkinter import ttk, simpledialog
from PIL import Image,ImageTk
import cv2
#from persiantools.jdatetime import JalaliDate
import datetime
from fpdf import FPDF
from tkinter import scrolledtext
from xlsxwriter import Workbook

#==========================================colors=====================================================
cream='#dad7cd'
white='white'
blue_dark='#00264d'
blue_light='#b3f0ff'
green_fosfori='#4dff4d'
green2='#a3b18a' 
green4='#3a5a40'
green5='#344e41'

#===============================================================Main Root SEtting====================================
mroot = Tk()
mroot.title("VEBER Electrical Calculation")
mroot.geometry("1050x450+100+50")
mroot.minsize(width = None,height=None)
mroot.resizable(width=True,height=True)
mroot.grid_propagate(False)
mroot.state("zoomed")
mroot.configure(bg=blue_dark)
#===================================================def==============================================================
def donothing():
   filewin = Toplevel(mroot)
   button = Button(filewin, text="Do nothing button")
   button.grid(row=0,column=0)
def space(frame,x,y,w,bg):
    space=ttk.Label(frame, text=" ",width=w,background=bg).grid(row=x,column=y,rowspan=1,columnspan=1,padx=3,pady=3,sticky='SNEW')

space(mroot,1,1,90,blue_dark)
language = "Per.Language"

def lang():
    global language
    font_name="B Titr" #Koodak, B Nazanin, B Titr,MRT_Lebanon
    font_size=11
    font_type="bold"
    if language == "Per.Language":
        #Fmb0.config(text="En.Language")
        mbutton1.config(text="نرم افزار محاسبات تابلو برق",width=32,height=2,font=(font_name, font_size, font_type))
        mbutton2.config(text="محاسبات زمان پشتیبانی یو پی اس",width=32,height=2,font=(font_name, font_size, font_type))
        mbutton3.config(text="محاسبات آمپراژ باتری یو پی اس",width=32,height=2,font=(font_name, font_size, font_type))
        mbutton4.config(text=" محاسبات دیزل ژنراتور",width=32,height=2,font=(font_name, font_size, font_type))
        mbutton5.config(text="محاسبات ترانسفورماتور",width=32,height=2,font=(font_name, font_size, font_type))
        company_lblm.config(text="گروه نرم افزاری مانی نیروی البرز",font=(font_name, 16, font_type))
        language = "En.Language"
    else:
        font_name="Corbel" # Helvetica, Cambria Math, Corbel, Tahoma, Verdana,Times New Roman,Microsoft Sans Serif
        font_size=12
        font_type="bold"
        #Fmb0.config(text="Per.Language")
        mbutton1.config(text="PANEL CALCULATION SOFTWARE",width=38,height=3,font=(font_name, font_size, font_type))
        mbutton2.config(text="UPS SUPPORT TIME CALCULATION",width=38,height=3,font=(font_name, font_size, font_type))
        mbutton3.config(text=" UPS BATTERY AMPERE CALCULATION",width=38,height=3,font=(font_name, font_size, font_type))
        mbutton4.config(text="DIESEL GENERATOR CALCULATION",width=38,height=3,font=(font_name, font_size, font_type))
        mbutton5.config(text="TRANS CALCULATION",width=38,height=3,font=(font_name, font_size, font_type))
        company_lblm.config(text="Mani Niroo Company",font=(font_name, 20, font_type))
        language = "Per.Language"
   

def bar(framebar):
    menubar = Menu(framebar)
    filemenu = Menu(menubar, tearoff=0)
    filemenu.add_command(label="New", command=donothing)
    filemenu.add_command(label="Open", command=donothing)
    filemenu.add_command(label="Save", command=donothing)
    filemenu.add_command(label="Save as...", command=donothing)
    filemenu.add_command(label="Print", command=donothing)
    filemenu.add_command(label="Close", command=donothing)

    filemenu.add_separator()
    filemenu.add_command(label="Exit", command=framebar.quit)
    menubar.add_cascade(label="File", menu=filemenu)

    editmenu = Menu(menubar, tearoff=0)
    editmenu.add_command(label="Undo", command=donothing)
    editmenu.add_separator()
    editmenu.add_command(label="Cut", command=donothing)
    editmenu.add_command(label="Copy", command=donothing)
    editmenu.add_command(label="Paste", command=donothing)
    editmenu.add_command(label="Delete", command=donothing)
    editmenu.add_command(label="Change Language", command=lambda:lang())

    menubar.add_cascade(label="Edit", menu=editmenu)

    helpmenu = Menu(menubar, tearoff=0)
    helpmenu.add_command(label="Help Index", command=donothing)
    helpmenu.add_command(label="About...", command=donothing)
    menubar.add_cascade(label="Help", menu=helpmenu)

    framebar.config(menu=menubar)

bar(mroot)
#====================================MAIN BUTTONS======================================
#mb0=Button(mroot,text="Per.Language",width=10,command= lambda:lang())
#Fmb0.grid(row=0,column=1)
mbtn_bg=green4
mbutton1 = Button(mroot,width=38,height=3, text="PANEL CALCULATION Software",font=("Corbel", "12", "bold"),background=mbtn_bg,foreground=white, command= lambda:create_panel_project())
mbutton1.grid(row=0,column=0, pady=5,padx=5)
mbutton2 = Button(mroot,width=38,height=3, text="UPS SUPPORT TIME CALCULATION",font=("Corbel", "12", "bold"), background=mbtn_bg,foreground=white,command=lambda:ups_support_time())
mbutton2.grid(row=1,column=0, pady=5,padx=5)
mbutton3 = Button(mroot,width=38,height=3, text="UPS BATTERY AMPERE CALCULATION",font=("Corbel", "12", "bold"),relief=RIDGE ,background=mbtn_bg,foreground=white, command=lambda:ups_battery_ampere())
mbutton3.grid(row=2,column=0, pady=5,padx=5)
mbutton4 = Button(mroot,width=38,height=3, text="DIESEL GENERATOR CALCULATION",font=("Corbel", "12", "bold"),background=mbtn_bg,foreground=white)#, command=lambda:ups_battery_ampere())
mbutton4.grid(row=3,column=0, pady=5,padx=5)
mbutton5 = Button(mroot,width=38,height=3, text="TRANS CALCULATION",font=("Corbel", "12", "bold"),background=mbtn_bg,foreground=white)#, command=lambda:ups_battery_ampere())
mbutton5.grid(row=4,column=0, pady=5,padx=5)
company_lblm=tk.Label(mroot,text="Mani Niroo Company",justify='right', font=('Corbel','16','bold'),foreground=white,background=blue_dark)
company_lblm.grid(row=2,column=4, pady=5,padx=5)
logo_path = Image.open(r"C:\Users\Seydi\Documents\LEARN_PYTHON\Projects\my_calculation\Images\logo khali.png")
logo_path = logo_path.resize((100, 100))
logo_image = ImageTk.PhotoImage(logo_path)
lbl_logo = tk.Label(image=logo_image)
lbl_logo.image = logo_image
lbl_logo.grid(row=0,column=4,rowspan=3)
project_list=[]


#=============================================def ups_battery_ampere =====================================================================
def ups_battery_ampere() :
    ups_name = (simpledialog.askstring("New Tab", "Enter the name for the UPS")).upper()
    ups_ampere_root = Toplevel()
    ups_ampere_root.title(f"CALCULATE OF {ups_name}")
    ups_ampere_root.geometry("350x300+50+50")    # By default, it is kept as the geometry of the main window, but you can change it.
    ups_ampere_root.wm_minsize(350,300)
    ups_ampere_root.wm_maxsize(350,300)
    top1=Frame(ups_ampere_root,bg=cream)           # پنجره top1
    top1.config(bd=4,height=5)
    top1.grid(row=0,column=0,rowspan=1,columnspan=1,sticky='snew')
    subroot1=Frame(ups_ampere_root,bg=green2) 
    subroot1.config(bd=4,height=5)
    subroot1.grid(row=1,column=0,rowspan=1,columnspan=1,sticky='snew')
    subroot1.columnconfigure(20, weight=1)
    subroot1.rowconfigure(0, weight=1)
    b=("Times", "16", "bold italic")
    lbl_title1=ttk.Label(top1,text=f"UPS BATTERY AMPERE: {ups_name}",font=b,justify='left',background=cream,width=30,anchor=CENTER)
    lbl_title1.grid(row=0,column=0,padx=5,pady=5,rowspan=1,columnspan=1,sticky=W)
    lbl_load=ttk.Label(subroot1, text="Load(KW) :",background='white',width=22)
    load_var=tk.DoubleVar()
    ent_load=ttk.Entry(subroot1,textvariable=load_var,background='white',width=8)

    lbl_time=ttk.Label(subroot1, text="Support Time(min) :",background='white',width=22)
    time_var=tk.IntVar()
    ent_time=ttk.Entry(subroot1,textvariable=time_var,background='white',width=8)
    lbl_quantity=ttk.Label(subroot1, text="Battery Quantity :",background='white',width=22)
    quantity_var=tk.IntVar()
    ent_quantity=ttk.Entry(subroot1,textvariable=quantity_var,background='white',width=8)
    lbl_eff=ttk.Label(subroot1, text="Battery Eff.(%) :",background='white',width=22)
    eff_var=tk.IntVar()
    ent_eff=ttk.Combobox(subroot1,textvariable=eff_var,background='white',width=6)
    ent_eff.config(value=(40,50,60,70,80,90,100),state='readonly')
    ent_eff.current(3)
    btn_calculate=Button(subroot1,text="Calculate",width=18, background=green_fosfori,command=lambda:Calculate1())
    lbl_load.grid(row=1,column=0,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    ent_load.grid(row=1,column=1,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    lbl_time.grid(row=2,column=0,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    ent_time.grid(row=2,column=1,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    lbl_quantity.grid(row=3,column=0,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    ent_quantity.grid(row=3,column=1,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    lbl_eff.grid(row=4,column=0,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    ent_eff.grid(row=4,column=1,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    btn_calculate.grid(row=5,column=0,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    lbl_load_kva1=ttk.Label(subroot1, text="Load(KVA) :",background='white',width=22)
    lbl_load_kva1.grid(row=6,column=0,rowspan=1,columnspan=2,padx=5,pady=5,sticky=W)
    lbl_battery_voltage=ttk.Label(subroot1, text="Battery Voltage(V) :",background='white',width=22)
    lbl_battery_voltage.grid(row=7,column=0,rowspan=1,columnspan=2,padx=5,pady=5,sticky=W)
    lbl_battery_amp=ttk.Label(subroot1,text="Battery Ampere(A) :",justify='left',background='white',width=22)
    lbl_battery_amp.grid(row=8,column=0,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    cal_load_kva1=ttk.Label(subroot1,text="",background=cream,width=8)
    cal_load_kva1.grid(row=6,column=1,rowspan=1,columnspan=2,padx=5,pady=5,sticky=W)
    cal_battery_voltage1=ttk.Label(subroot1,text="",background=cream,width=8)
    cal_battery_voltage1.grid(row=7,column=1,rowspan=1,columnspan=2,padx=5,pady=5,sticky=W)
    cal_battery_amp1=ttk.Label(subroot1,text="",justify='left',background=cream,width=8,anchor=W)
    cal_battery_amp1.grid(row=8,column=1,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    bar(ups_ampere_root)

    def Calculate1():
        ups_data1=[]
        load1=float(ent_load.get())
        time1=int(ent_time.get())
        eff1=int(eff_var.get())
        quantity1=int(ent_quantity.get())
        battery_voltage_var1=quantity1*12
        load_kva1=load1/(0.8)
        battery_amp1=int((load1*1000*(time1/60))/(battery_voltage_var1*(eff1/100)))
        cal_load_kva1.config(text=load_kva1,background=green_fosfori)
        cal_battery_voltage1.config(text=battery_voltage_var1,background=green_fosfori)
        cal_battery_amp1.config(text=battery_amp1,background=green_fosfori)

        ups_data1.append([ups_name,load1,load_kva1,battery_amp1,eff1,quantity1,battery_voltage_var1,battery_amp1])
        df_ups_support_time=pd.DataFrame(ups_data1,columns=['UPS Name','Load(KW)','Load(KVA)','Backup Time(min)','Battery effecienty(%)','Battery Quantity','Battery Voltage(V)','Batterry Ampere(A)'])
        print (df_ups_support_time)
        writer = pd.ExcelWriter("UPS Battery Ampere.xlsx", engine="xlsxwriter")
        df_ups_support_time.to_excel(writer, sheet_name=(ups_name), startrow=0, header=True,startcol=0,float_format="%.2f")
        writer.close()

#=====================================================def ups_support_time=======================================
def ups_support_time():
    ups_name = (simpledialog.askstring("New Tab", "Enter the name for the UPS")).upper()
    
    ups_root2= Toplevel()
    ups_root2.title(f"CALCULATE OF {ups_name}")
    ups_root2.geometry("350x300+500+50")# By default, it is kept as the geometry of the main window, but you can change it.
    ups_root2.wm_minsize(350,300)
    ups_root2.wm_maxsize(350,300)
    top2=Frame(ups_root2,bg=cream)           # پنجره top2
    top2.config(bd=4,height=5)
    top2.grid(row=0,column=1,rowspan=1,columnspan=1,sticky='snew')

    subroot2=Frame(ups_root2,bg=green2)           # پنجره subroot2
    subroot2.config(bd=4,height=5)
    subroot2.grid(row=1,column=1,rowspan=1,columnspan=1,sticky='SNEW')
    
    subroot2.columnconfigure(20, weight=1)
    subroot2.rowconfigure(0, weight=1)
    b=("Times", "16", "bold italic")

    lbl_title2=ttk.Label(top2,text=f"UPS SUPPORT TIME: {ups_name}",font=b,justify='left',background=cream,width=30,anchor=CENTER)
    lbl_title2.grid(row=0,column=0,padx=5,pady=5,rowspan=1,columnspan=1,sticky=W)

    lbl_load2=ttk.Label(subroot2, text="Load(KW) :",background='white',width=22)
    load_var2=tk.DoubleVar()
    ent_load2=ttk.Entry(subroot2,textvariable=load_var2,background='white',width=8)

    battery_amp_var2=tk.IntVar()
    lbl_battery_amp2=ttk.Label(subroot2,text="Battery Ampere(A) :",justify='left',background='white',width=22)
    ent_battery_amp2=ttk.Entry(subroot2,textvariable=battery_amp_var2,background='white',width=8)

    lbl_quantity2=ttk.Label(subroot2, text="Battery Quantity :",background='white',width=22)
    quantity_var2=tk.IntVar()
    ent_quantity2=ttk.Entry(subroot2,textvariable=quantity_var2,background='white',width=8)

    lbl_eff2=ttk.Label(subroot2, text="Battery Eff.(%) :",background='white',width=22)
    eff_var2=tk.IntVar()
    ent_eff2=ttk.Combobox(subroot2,textvariable=eff_var2,background='white',width=6)
    ent_eff2.config(value=(40,50,60,70,80,90,100),state='readonly')
    ent_eff2.current(3)

    btn_calculate2=Button(subroot2,text="Calculate",width=18, background=green_fosfori,command=lambda:Calculate2())

    lbl_load2.grid(row=1,column=0,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    ent_load2.grid(row=1,column=1,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
   
    lbl_battery_amp2.grid(row=2,column=0,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    ent_battery_amp2.grid(row=2,column=1,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)

    lbl_quantity2.grid(row=3,column=0,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    ent_quantity2.grid(row=3,column=1,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)

    lbl_eff2.grid(row=4,column=0,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    ent_eff2.grid(row=4,column=1,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    btn_calculate2.grid(row=5,column=0,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    
    
    lbl_load_kva2=ttk.Label(subroot2, text="Load(KVA) :",background='white',width=22)
    lbl_load_kva2.grid(row=6,column=0,rowspan=1,columnspan=2,padx=5,pady=5,sticky=W)
    

    lbl_battery_voltage2=ttk.Label(subroot2, text="Battery Voltage(V) :",background='white',width=22)
    lbl_battery_voltage2.grid(row=7,column=0,rowspan=1,columnspan=2,padx=5,pady=5,sticky=W)

    lbl_time2=ttk.Label(subroot2, text="Support Time(min) :",background='white',width=22)
    lbl_time2.grid(row=8,column=0,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)
    
    cal_load_kva2=ttk.Label(subroot2,text="",background=cream,width=8)
    cal_load_kva2.grid(row=6,column=1,rowspan=1,columnspan=2,padx=5,pady=5,sticky=W)
    cal_battery_voltage2=ttk.Label(subroot2,text="",background=cream,width=8)
    cal_battery_voltage2.grid(row=7,column=1,rowspan=1,columnspan=2,padx=5,pady=5,sticky=W)
    cal_time2=ttk.Label(subroot2,text="",justify='left',background=cream,width=8,anchor=W)
    cal_time2.grid(row=8,column=1,rowspan=1,columnspan=1,padx=5,pady=5,sticky=W)

    bar(ups_root2)
    def Calculate2():
        ups_data2=[]
        load2=float(ent_load2.get())
        battery_amp2=int(ent_battery_amp2.get())
        load_kva2=load2/(0.8)
        eff2=int(eff_var2.get())
        quantity2=int(ent_quantity2.get())
        battery_voltage_var2=quantity2*12   
        time2_var=int(60*((battery_voltage_var2*(eff2/100)*battery_amp2)/(load2*1000))) 
        
        cal_load_kva2.config(text=load_kva2,background=green_fosfori)
        cal_battery_voltage2.config(text=battery_voltage_var2,background=green_fosfori)
        cal_time2.config(text=time2_var,background=green_fosfori)
        
        ups_data2.append([ups_name,load2,load_kva2,battery_amp2,eff2,quantity2,battery_voltage_var2,time2_var])
        df_ups_support_time=pd.DataFrame(ups_data2,columns=['UPS Name','Load(KW)','Load(KVA)','Batterry Ampere(A)','Battery effecienty(%)','Battery Quantity','Battery Voltage(V)','Backup Time(min)'])
        print (df_ups_support_time)
        writer = pd.ExcelWriter("UPS Support Time.xlsx", engine="xlsxwriter")
        df_ups_support_time.to_excel(writer, sheet_name=(ups_name), startrow=0, header=True,startcol=0,float_format="%.2f")
        writer.close()
#======================================================  def create_panel_project ===============================================     
def create_panel_project():
    rootpanel=Tk()
    project =Project(rootpanel)
    rootpanel.mainloop()
def space(frame,x,y,bg):
    space=ttk.Label(frame, text=" ",width=4,background=bg).grid(row=x,column=y,rowspan=1,columnspan=1,padx=3,pady=3,sticky='SNEW')
    
def holder(frame,x,y,w):
    holder_lbl=ttk.Label(frame, text="", width=w,background=cream).grid(row=x,column=y,rowspan=1,columnspan=1,padx=3,pady=3,sticky='W')

class Project:
    def __init__(self,rootpanel):
      
        project_name = simpledialog.askstring("New Tab", "Enter the name for the Project")
        while  project_name=='':
            ValueError
            messagebox.showerror("Error", "Please Insert a Name for Project")
            project_name = simpledialog.askstring("New Tab", "Enter the name for the Project")   
              
        project_name=project_name.upper()
        self.rootpanel = rootpanel
        self.rootpanel.title(f"{project_name} Project")
        self.rootpanel.geometry("1300x850+100+50")
        self.rootpanel.state("zoomed") 
       
        
        self.header=Frame(rootpanel,bg=blue_dark,height=200,width=rootpanel.winfo_screenwidth())
        self.header.grid(row=0,column=0,rowspan=1,columnspan=1,sticky='snew')
        
        self.under_header=Frame(rootpanel,bg=green2,height=2,width=rootpanel.winfo_screenwidth())
        self.under_header.grid(row=1,column=0,rowspan=1,columnspan=1,sticky='snew')
        
        self.notebook = ttk.Notebook(rootpanel,height=620,width=rootpanel.winfo_screenwidth())# ایجاد نوت‌بوک برای تب‌ها
        self.notebook.grid(row=3,column=0,rowspan=1,columnspan=1,sticky='snew') 
                 
        self.tab_control = ttk.Notebook(rootpanel,height=620,width=rootpanel.winfo_screenwidth())
        self.panel_list=[] # لیست برای ذخیره نام تب‌ها
        
        
        a=("Helvetica", "10", "bold")
        c=("Literal", "12", "bold")
    
        self.space_header=tk.Label(self.header,text="",width=26,background=blue_dark)
        self_add_button = tk.Button(self.header, text="Add Panel",width=24,font=a,background=green4,foreground=white,command=self.add_panel)
        self.delete_button = tk.Button(self.header, text="Delete Panel",width=24,justify='left',font=a,background=green4,foreground=white,command=self.delete_active_tab)
        self.excell_button = tk.Button(self.header, text="To Excell",width=24,justify='left',font=a,background=green4,foreground=white,command=self.to_excell)
        self.pdf_button=tk.Button(self.header, text="To PDF",width=24,justify='left',font=a,background=green4,foreground=white)#,command=self.to_PDF)
        self.company_lbl=tk.Label(self.header,text="گروه نرم افزاری مانی نیروی البرز",justify='left', font=('Vazir','16','bold'),foreground="white",background=blue_dark)
        
        self_add_button.grid(row=0,column=1,padx=5,pady=5,sticky='snew')
        self.delete_button.grid(row=0,column=2,padx=5,pady=5,sticky='snew')
        self.excell_button.grid(row=0,column=3,padx=5,pady=5,sticky='snew')
        self.pdf_button.grid(row=0,column=4,padx=5,pady=5,sticky='snew')
        self.space_header.grid(row=0,column=5,padx=5,pady=5,sticky='snew')
        self.company_lbl.grid(row=0,column=6,padx=5,pady=5,sticky='snew')

        bar(rootpanel)
        project_list.append(project_name)
        print(project_list)
        
        
        ptab = ttk.Frame(self.notebook,width=self.rootpanel.winfo_screenwidth(),height=600)
        self.notebook.add(ptab, text='Project Info('+project_name+')')
        self.notebook.select(ptab)
        

        project_top0=Frame(ptab,bg=blue_dark,border=2,height=100,width=1500) # پنجره panel_top0
        project_top0.grid_propagate(False)
        project_top0.grid(row=0,column=0,rowspan=1,columnspan=1,sticky='snew')
        
        
        project_top1=Frame(ptab,bg=blue_dark,border=2,height=550,width=1500)           # پنجره panel_top1
        project_top1.grid_propagate(False)
        project_top1.grid(row=1,column=0,rowspan=1,columnspan=1,sticky='sn')

        
        print(self.notebook.tab)
        
        
        
        lbl_project_prop=ttk.Label(project_top0, text="Project Information:",font=('Literal','22'),justify='center',background=blue_dark,foreground=white,width=22)
        
        lbl_project_name=ttk.Label(project_top1, text="Project Name:",font=c,justify='left',background='white',width=18)
        ent_project_name=ttk.Label(project_top1, text=project_name,font=c,justify='left',background='white',width=24)
       
        lbl_project_address=ttk.Label(project_top1, text="Project Address:",font=c,justify='left',background='white',width=18)
        ent_project_address=ttk.Entry(project_top1, text="",font=c,justify='left',background='white',width=48)
        
        lbl_client_name=ttk.Label(project_top1, text="Client Name:",font=c,justify='left',background='white',width=18)
        ent_client_name=ttk.Entry(project_top1, text='',font=c,justify='left',background='white',width=24)
       
        lbl_client_address=ttk.Label(project_top1, text="Client Address:",font=c,justify='center',background='white',width=18)
        ent_client_address=ttk.Entry(project_top1, text="",font=c,justify='left',background='white',width=24)
        
        lbl_project_designer=ttk.Label(project_top1, text="Designer Name:",font=c,justify='center',background='white',width=18)
        ent_project_designer=ttk.Entry(project_top1, text="",font=c,justify='left',background='white',width=24)
        
        lbl_designer_contact=ttk.Label(project_top1, text="Designer Contact:",font=c,justify='center',background='white',width=18)
        ent_designer_contact=ttk.Entry(project_top1, text="",font=c,justify='left',background='white',width=24)

        
        lbl_project_prop.grid(row=0,column=1,rowspan=2,columnspan=1,padx=10,pady=10)
        
        
        lbl_space=ttk.Label(project_top1, text="",justify='center',background=blue_dark,width=50)
        lbl_space.grid(row=1,column=0,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')
        
        lbl_space2=ttk.Label(project_top1, text="",justify='center',background=blue_dark,width=50)
        lbl_space2.grid(row=1,column=6,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')
        
        lbl_space3=ttk.Label(project_top1, text="",justify='center',background=blue_dark,width=50)
        lbl_space3.grid(row=7,column=1,rowspan=24,columnspan=1,padx=10,pady=10,sticky='snew')



        lbl_project_name.grid(row=1,column=1,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')
        ent_project_name.grid(row=1,column=2,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')

        lbl_project_address.grid(row=2,column=1,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')
        ent_project_address.grid(row=2,column=2,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')
        

        lbl_client_name.grid(row=3,column=1,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')
        ent_client_name.grid(row=3,column=2,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')
        
        lbl_client_address.grid(row=4,column=1,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')
        ent_client_address.grid(row=4,column=2,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')
        
        lbl_project_designer.grid(row=5,column=1,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')
        ent_project_designer.grid(row=5,column=2,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')
      
        lbl_designer_contact.grid(row=6,column=1,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')
        ent_designer_contact.grid(row=6,column=2,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')
        
    
        self.panel_list.append('Project Info('+project_name+')')    
        
        
        
        
        
        

        
        
    def add_panel(self):
        panel_name = simpledialog.askstring("New Tab", "Enter the name for the new Panel")
        tab = ttk.Frame(self.notebook,width=self.rootpanel.winfo_screenwidth(),height=600)
        panel_name=panel_name.upper()
        self.notebook.add(tab, text=panel_name)
        self.notebook.select(tab)
        print(self.notebook.tab)
        
        self.panel_list.append(panel_name)
        print(self.panel_list)
        
        
        
        
        
        global breakers,feeder_types
        breakers=(6,10,16,20,25,32,40,50,63,80,100,125,160,200,250,320,400,630,800,1000,1250,1600,2000)
        feeder_types= ("Lighting","Socket","Equipment","Motor(1P-DOL)","Motor(3P-DOL)","Motor(3P-YD)","Panel")
        feeders_data = []
        panel_data=[]
        total_power_var= 0
        current_list = [0, 0, 0]
        demand_current=0
        rows=[]
        rows2=[]
        rows3=[]
        current_outs=[]
        cable_out=[]
        phase_n_out=[]
        delta_v_out=[]
        breaker_out=[]
        lbl_f_current=0
        current_val=0
        cable=0
        f_phase_name_var="_"
        
        panel_top1=Frame(tab,bg=blue_dark,border=2)           # پنجره panel_top1
        panel_top1.grid(row=0,column=0,rowspan=1,columnspan=1,sticky='SNEW')

        panel_top2=Frame(tab,bg=blue_dark,border=2)           # پنجره panel_top2
        panel_top2.grid(row=1,column=0,rowspan=1,columnspan=1,sticky='SNEW')

        panel_top3=Frame(tab,bg=white,border=2)           # پنجره panel_top3
        panel_top3.grid(row=2,column=0,rowspan=1,columnspan=1,sticky='SNEW')

        panel_top4=Frame(tab,bg=cream,border=2)            # پنجره panel_top4
        panel_top4.grid(row=3,column=0,rowspan=1,columnspan=1,sticky='SNEW')

        frame2 = ttk.Frame(tab,width=tab.winfo_screenheight(),height=420)
        frame2.grid(row=4, column=0,rowspan=1,columnspan=1, sticky="nsew")

        # =================================Step 4: Create a Canvas and Scrollbar========================================
        canvas = tk.Canvas(frame2)
        scrollbar = ttk.Scrollbar(frame2, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set,width=1300,height=420)
        # Step 5: Create a Frame for Scrollable Content
        panel_top5 = ttk.Frame(canvas)
        # Step 6: Configure the Canvas and Scrollable Content Frame
        panel_top5.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        # Step 8: Create Window Resizing Configuration
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(4, weight=1)
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)
        panel_top1.columnconfigure(20, weight=1)
        panel_top2.columnconfigure(20, weight=1)

        ###################################################################
        # Step 9: Pack Widgets onto the Window
        canvas.create_window((0, 0), window=panel_top5, anchor="nw")
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        # Step 10: Bind the Canvas to Mousewheel Events
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        
        #=========================================def space============================================
        

        #======================================================DEF CAL===================================================
        def calculate():
            file_path = "data.xlsx"
            # بارگذاری دو شیت از اکسل
            sheets = {
                "In Air": pd.read_excel(file_path, sheet_name="In Air"),
                "In Ground": pd.read_excel(file_path, sheet_name="In Ground")
            }
            
            print('________*_________')
            m3_yd=pd.read_excel('data.xlsx', sheet_name='m3_yd',converters={'POWER':float,'BREAKER':int,'SETTING':int,'BMETAL':str,'CONTACTOR':int,'CABLE':float})
            m3_dol=pd.read_excel('data.xlsx', sheet_name='m3_dol',converters={'POWER':float,'BREAKER':int,'SETTING':int,'BMETAL':str,'CONTACTOR':int,'CABLE':float})
            m1_dol=pd.read_excel('data.xlsx', sheet_name='m1_dol',converters={'POWER':float,'BREAKER':int,'SETTING':int,'BMETAL':str,'CONTACTOR':int,'CABLE':float})
            current_table=pd.read_excel('data.xlsx', sheet_name='current_table')
            lighting_table=pd.read_excel('data.xlsx', sheet_name='lighting_table')
            socket_table=pd.read_excel('data.xlsx', sheet_name='socket_table')
            breaker=pd.read_excel('data.xlsx', sheet_name='breaker')
            panel_table=pd.read_excel('data.xlsx', sheet_name='panel_table')
            #panel_name_var=panel_name_input.get()
            upstream_panel_name=upstream_panel.get()
            panel_phase_name=panel_phase_input.get()
            panel_cable_len=self.main_cable_len_var.get()
            panel_d_f=float(self.panel_d_f_var.get())
            temp=temp_intry.get()
            print(temp)
            print("1")
            
            installation_type=  instalation_input.get()
            print(installation_type)
            print("2")
            
            insulation= c_insulation_input.get()
            print(insulation)
            print("3")
            
            max_volage_drop=float(self.max_volage_drop_var.get())

            if installation_type and insulation and temp:
                kf = sheets.get(installation_type)
            try:
                k = kf[kf["TEMP"] == int(temp)][insulation].values[0]
                k=float(k)
                
            except (IndexError, KeyError):
                messagebox.showerror("خطا", "مقدار مورد نظر یافت نشد.")
            #else:
                #messagebox.showwarning("توجه", "لطفاً همه‌ی مقادیر را انتخاب کنید.")

        
            if panel_phase_name== "RST":
                panel_phase=3
            else:
                panel_phase=1

            panel_d_f=float(panel_d_f_input.get())
            try:
                if 1>=panel_d_f or 0<panel_d_f :
                    pass
            except ValueError:
                messagebox.showerror("Error", "Please Insert a Correct DEMAND FACTOR (0<D.F=<1)")
            try:
                main_cable_l=int(panel_cable_len)
            except ValueError:
                messagebox.showerror("Error", "Please Insert :Main Cable Lentgh(m) ")
 

            print (k)

            while  rows2:
                last_row2 = rows2.pop()   
                for widget in last_row2:
                    widget.destroy()
            total_power_var=0
            demand_current=0
            cable=""
            cb=""
            panel_data.clear()
            current_outs.clear()
            cable_out.clear()
            phase_n_out.clear()
            delta_v_out.clear()
            breaker_out.clear()
            feeders_data.clear()

            max_current = 0
            demand_current=0
            demand_power=0

            current_list=[0,0,0]

            for row in rows:

                f_type_var=str(row[1].get())
                f_power_var=row[2].get()
                f_phase_var = row[3].get()
                f_pf_var =row[4].get()
                f_cable_len_var=row[5].get()
                f_number =row[0].cget("text")


                if f_power_var == "" or f_pf_var == "" or f_cable_len_var =="" :
                    messagebox.showerror("Error", "Please Insert a Value")
                    return
                try:
                    f_power_var = float(f_power_var)
                    f_cable_len_var=int(f_cable_len_var)
                    f_pf_var=( float(f_pf_var))
                    if 1<f_pf_var or 0>=f_pf_var :
                        messagebox.showerror("Error","Please Insert a Correct POWER FACTOR (0<P.F=<1)")
                    else:
                        pass  
                    
                except ValueError:
                    messagebox.showerror("Error", "Please Insert a Correct Value")
                    break
                
                total_power_var += f_power_var
                current_val="{:.1f}".format((float(f_power_var))*1000/(int(f_phase_var)*float(f_pf_var)*230))
                f_phase_var=int(f_phase_var)
                panel_phase_name=panel_phase_input.get()
                if panel_phase_name=="RST":
                    if f_phase_var == 1 :
                        min_index = current_list.index(min(current_list))
                        current_list[min_index] +=  float(current_val)
                        ph_name_help = min_index + 1
                        f_phase_name_var = "R" if ph_name_help == 1 else "S" if ph_name_help == 2 else "T"

                    else:
                        current_list[0] += float(current_val)
                        current_list[1] += float(current_val)
                        current_list[2] += float(current_val)
                        f_phase_name_var = "RST"
                else:
                    f_phase_name_var=panel_phase_name
                    if f_phase_name_var =="R":
                        current_list[0] += float(current_val)
                    elif f_phase_name_var =="S" :
                        current_list[1] += float(current_val)
                    elif f_phase_name_var =="T" :
                        current_list[2] += float(current_val)
                kc=(float(1.25))
                f_current=float(current_val)
                current_val_e=f_current*kc

                fi=int(0)

                if f_type_var=="Lighting" and f_phase_var==1  :
                    filtered_f_cb = breaker[breaker['c_breaker'] >= current_val_e]
                    cb = filtered_f_cb.iloc[0][1]
                    filtered_f_cable = lighting_table[lighting_table['1PHASE_A'] >= cb]
                    cable = filtered_f_cable.iloc[fi][5]
                    bmetal="-"
                    setting="-"
                    contactor="-"

                elif f_type_var=="Lighting" and f_phase_var==3  :
                    filtered_f_cb = breaker[breaker['c_breaker'] >= current_val_e]
                    cb = filtered_f_cb.iloc[0][1]
                    filtered_f_cable = lighting_table[lighting_table['3PHASE_A'] >= cb]
                    cable = filtered_f_cable.iloc[fi][5]
                    bmetal="-"
                    setting="-"
                    contactor="-"



                elif f_type_var=="Socket" or f_type_var=="Equipment" and f_phase_var==1 :
                    filtered_f_cb = breaker[breaker['c_breaker'] >= current_val_e] 
                    cb = filtered_f_cb.iloc[0][1]
                    filtered_f_cable = socket_table[socket_table['1PHASE_A'] >= cb]
                    cable = filtered_f_cable.iloc[fi][5]
                    bmetal="-"
                    setting="-"
                    contactor="-"

                elif f_type_var=="Socket" or f_type_var=="Equipment" and f_phase_var==3 :
                    filtered_f_cb = breaker[breaker['c_breaker'] >= current_val_e] 
                    cb = filtered_f_cb.iloc[0][1]
                    filtered_f_cable = socket_table[socket_table['3PHASE_A'] >= cb]
                    cable = filtered_f_cable.iloc[fi][5]
                    bmetal="-"
                    setting="-"
                    contactor="-"



                elif f_type_var=="Panel"and f_phase_var==3 :
                    filtered_f_cb = breaker[breaker['c_breaker'] >= (current_val_e) ] 
                    cb = filtered_f_cb.iloc[0][1]
                    filtered_f_cable =panel_table [panel_table['3PHASE_A'] >= cb]
                    cable = filtered_f_cable.iloc[fi][5]
                    bmetal="-"
                    setting="-"
                    contactor="-"

                elif f_type_var=="Panel"and f_phase_var==1 :
                    filtered_f_cb = breaker[breaker['c_breaker'] >= (current_val_e)] 
                    cb = filtered_f_cb.iloc[0][1]   
                    filtered_f_cable =panel_table [panel_table['1PHASE_A'] >= cb]
                    cable = filtered_f_cable.iloc[fi][5]
                    bmetal="-"
                    setting="-"
                    contactor="-"

                elif f_type_var=="Motor(1P-DOL)" :
                    filtered_f_cb = m1_dol[m1_dol['POWER'] >= f_power_var] 
                    cb = filtered_f_cb.iloc[0][1]  
                    filtered_f_cable = m1_dol[m1_dol['POWER'] >= (f_power_var)]
                    cable = filtered_f_cable.iloc[fi][5]
                    bmetal=filtered_f_cable.iloc[0][3]
                    setting=filtered_f_cable.iloc[0][2]
                    contactor=filtered_f_cable.iloc[0][4]

                elif f_type_var=="Motor(3P-DOL)" :
                    filtered_f_cb = m3_dol[m3_dol['POWER'] >= (f_power_var)] 
                    cb = filtered_f_cb.iloc[0][1]  
                    filtered_f_cable = m3_dol[m3_dol['POWER'] >= (f_power_var)]
                    cable = filtered_f_cable.iloc[fi][5]
                    bmetal=filtered_f_cable.iloc[0][3]
                    setting=filtered_f_cable.iloc[0][2]
                    contactor=filtered_f_cable.iloc[0][4]

                elif f_type_var=="Motor(3P-YD)" and f_phase_var==3:
                    filtered_f_cb = m3_yd[m3_yd['POWER'] >= (f_power_var)] 
                    cb = filtered_f_cb.iloc[0][1]  
                    filtered_f_cable = m3_yd[m3_yd['POWER'] >= (f_power_var)]
                    cable = filtered_f_cable.iloc[fi][5]
                    bmetal=filtered_f_cable.iloc[0][3]
                    setting=filtered_f_cable.iloc[0][2]
                    contactor=filtered_f_cable.iloc[0][4]


                if f_phase_var == 1 :
                    f_delta_v =(float((float(f_power_var))*1000*f_cable_len_var*2)/(56*cable*230))

                else:
                    f_delta_v =(float((float(f_power_var))*1000*f_cable_len_var)/(56*cable*400))

                while f_delta_v >4:
                    fi=int(fi+1)
                    cable = filtered_f_cable.iloc[fi][5]
                    if f_phase_var == 1 :
                        f_delta_v =(float((float(f_power_var))*1000*f_cable_len_var*2)/(56*cable*230))

                    else:
                        f_delta_v =(float((float(f_power_var))*1000*f_cable_len_var)/(56*cable*400))



                lbl_f_current_c = ttk.Label(panel_top5,text=current_val, width=12,background=green_fosfori,justify='center')
                lbl_f_current_c.grid(row=len(rows2), column=6,padx=3,pady=2,sticky='NW')


                lbl_f_cable_c = ttk.Label(panel_top5,text=cable, width=12,background=green_fosfori,justify='center')
                lbl_f_cable_c.grid(row=len(rows2), column=7,padx=3,pady=2,sticky='NW')   

                lbl_f_phase_name_c = ttk.Label(panel_top5,text=f_phase_name_var, width=12,background=green_fosfori,justify='center')
                lbl_f_phase_name_c.grid(row=len(rows2), column=8,padx=2,pady=3,sticky='NW')

                lbl_f_delta_v_c = ttk.Label(panel_top5,text="{:.2f}".format(f_delta_v), width=12,background=green_fosfori,justify='center')
                lbl_f_delta_v_c.grid(row=len(rows2), column=9,padx=3,pady=2,sticky='NW') 

                lbl_f_breaker_c = ttk.Label(panel_top5,text=cb, width=12,background=green_fosfori,justify='center')
                lbl_f_breaker_c.grid(row=len(rows2), column=10,padx=3,pady=2,sticky='NW')  
                lbl_f_name_c = ttk.Label(panel_top5,text=(f"F{len(rows2)+1}"),width=12,background=green_fosfori)
                lbl_f_name_c.grid(row=len(rows2), column=11,padx=3,pady=2,sticky='NW')
                lbl_f_bmetal_c=ttk.Label(panel_top5,text=bmetal,width=12,background=green_fosfori)
                lbl_f_bmetal_c.grid(row=len(rows2), column=12,padx=3,pady=2,sticky='NW')
                lbl_f_setting_c=ttk.Label(panel_top5,text=setting,width=12,background=green_fosfori)
                lbl_f_setting_c.grid(row=len(rows2), column=13,padx=3,pady=2,sticky='NW')
                lbl_f_contactor_c=ttk.Label(panel_top5,text=contactor,width=12,background=green_fosfori)
                lbl_f_contactor_c.grid(row=len(rows2), column=14,padx=3,pady=2,sticky='NW')



                feeders_data.append([f_number,f_type_var,f_power_var,f_phase_var,f_pf_var,f_cable_len_var,current_val,cable,f_phase_name_var,f_delta_v,cb,(f"F{len(rows2)+1}"),bmetal,setting,contactor])
                rows2.append([lbl_f_current_c,lbl_f_cable_c,lbl_f_phase_name_c,lbl_f_delta_v_c,lbl_f_breaker_c,lbl_f_name_c,lbl_f_bmetal_c,lbl_f_setting_c,lbl_f_contactor_c])
                

            max_current = float(max(current_list))
            demand_current=max_current*(panel_d_f)
            demand_power=total_power_var*(panel_d_f)
            derated_current=(demand_current)/(k)

            if panel_phase_name=="RST"  :
                panel_pf_var=demand_power*1000/(230*3*demand_current)
            else:
                panel_pf_var=demand_power*1000/(230*demand_current)


            filtered_breaker = breaker[breaker['c_breaker'] >= demand_current*1.25]
            main_breaker = filtered_breaker.iloc[0][1]
            if main_breaker<=16:
                main_breaker = filtered_breaker.iloc[2][1]



            filtered_upstream_cb = breaker[breaker['c_breaker'] == main_breaker]
            upstream_cb =filtered_upstream_cb.iloc[0][2]
            i=0
            if panel_phase==3 and installation_type=="In Air" and upstream_cb>= derated_current :
                filtered_current_table = current_table[current_table['3PHASE_A'] >= upstream_cb]
                maine_cable_size = filtered_current_table.iloc[i][4]
            elif panel_phase==3 and installation_type=="In Air" and upstream_cb< derated_current :
                filtered_current_table = current_table[current_table['3PHASE_A'] >= derated_current]
                maine_cable_size = filtered_current_table.iloc[i][4]
            elif panel_phase==3 and installation_type=="In Ground" and upstream_cb< derated_current:
                filtered_current_table = current_table[current_table['3PHASE_G'] >= derated_current]
                maine_cable_size = filtered_current_table.iloc[i][4]
            elif panel_phase==3 and installation_type=="In Ground" and upstream_cb>= derated_current:
                filtered_current_table = current_table[current_table['3PHASE_G'] >= upstream_cb]
                maine_cable_size = filtered_current_table.iloc[i][4]

            elif panel_phase==1 and installation_type=="In Air" and upstream_cb< derated_current:
                filtered_current_table = current_table[current_table['1PHASE_A'] >= derated_current]
                maine_cable_size = filtered_current_table.iloc[i][4]
            elif panel_phase==1 and installation_type=="In Air" and upstream_cb>= derated_current:
                filtered_current_table = current_table[current_table['1PHASE_A'] >= upstream_cb]
                maine_cable_size = filtered_current_table.iloc[i][4]


            elif panel_phase==1 and installation_type=="In Ground" and upstream_cb< derated_current:
                filtered_current_table = current_table[current_table['1PHASE_G'] >= derated_current]
                maine_cable_size = filtered_current_table.iloc[i][4]
            elif panel_phase==1 and installation_type=="In Ground" and upstream_cb>= derated_current:
                filtered_current_table = current_table[current_table['1PHASE_G'] >= upstream_cb]
                maine_cable_size = (filtered_current_table.iloc[i][4]).get()

            def calc_delta_v():  
                if panel_phase == 1 :
                      return (float((float(demand_power))*1000*panel_cable_len*2)/(56*maine_cable_size*230))
                else:
                     return (float((float(demand_power))*1000*panel_cable_len)/(56*maine_cable_size*400))
            p_delta_v11 =calc_delta_v()
            p_delta_v=float(p_delta_v11)
            while p_delta_v >max_volage_drop:
                i=i+1
                maine_cable_size = filtered_current_table.iloc[i][4]
                p_delta_v=calc_delta_v()





            #total_power_var=tk.DoubleVar()    
            lbl_total_power.config(text="{:.2f}".format(total_power_var),justify='left',background=green_fosfori)
            lbl_panel_current=ttk.Label(panel_top2,text="{:.1f}".format(max_current),justify='left',background=green_fosfori,width=8)
            lbl_total_d_power=ttk.Label(panel_top2,text="{:.2f}".format(demand_power),justify='left',background=green_fosfori,width=8)
            lbl_panel_d_current=ttk.Label(panel_top2,text="{:.1f}".format(demand_current),justify='left',background=green_fosfori,width=8)
            lbl_panel_pf=ttk.Label(panel_top2,text="{:.2f}".format(panel_pf_var),justify='left',background=green_fosfori,width=8)
            #lbl_total_power.grid(row=0,column=2,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)
            lbl_panel_current.grid(row=1,column=2,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)
            lbl_total_d_power.grid(row=0,column=5,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)
            lbl_panel_d_current.grid(row=1,column=5,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)  
            lbl_panel_pf.grid(row=0,column=8,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W) 
            lbl_panel_breaker=ttk.Label(panel_top2,text=main_breaker,justify='left',background=green_fosfori,width=8)
            lbl_panel_breaker.grid(row=0,column=11,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)  
            lbl_panel_upstream_cb=ttk.Label(panel_top2,text=upstream_cb,justify='left',background=green_fosfori,width=8)
            lbl_panel_upstream_cb.grid(row=0,column=15,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)  
            lbl_panel_cable_size=ttk.Label(panel_top2,text=maine_cable_size,justify='left',background=green_fosfori,width=8)
            lbl_panel_cable_size.grid(row=1,column=8,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)
            lbl_panel_delta_v=ttk.Label(panel_top2,text="{:.2f}".format(p_delta_v),justify='left',background=green_fosfori,width=8)
            lbl_panel_delta_v.grid(row=1,column=11,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)
            fd_input = pd.DataFrame(feeders_data, columns=['NO','Feeder Type','Power (KW)','Phase', 'Power Factor', 'Cable Length (m)','Current (A)','Cable Size (mm2)','Phase Name','Delta V%','Breaker (A)','F.Name','Bmetal','Setting','Contactor'])   
            panel_data.append([panel_name,panel_phase,panel_d_f,"{:.2f}".format(total_power_var),"{:.2f}".format(demand_power),"{:.1f}".format(max_current),"{:.1f}".format(demand_current),upstream_panel_name,main_cable_l,temp,maine_cable_size,p_delta_v])
            pd_input = pd.DataFrame(panel_data, columns=['Panel Name','Panel Phase','P.D.F','C.Load (KW)','D.Load (KW)','C.Current (A)','D.Current (A)', 'Upstream P.', 'Cable Length (m)','Amb.TEMP','Main Cable (mm2)','Delta V'])
            print(pd_input)
            print(fd_input)
            
            #==================
            
        
            df1 = pd.DataFrame(panel_data)
            with pd.ExcelWriter('sample.xlsx') as writer:  
                df1.to_excel(writer, sheet_name='m1')
            
                with pd.ExcelWriter('sample.xlsx', engine='openpyxl', mode='a') as writer:  
                    pd_input.to_excel(writer, sheet_name=f'x{i}',startrow=0,header=True,startcol=0,float_format="%.2f")
                    fd_input.to_excel(writer, sheet_name=f'x{i}',startrow=7,header=True,startcol=0,float_format="%.2f")
            
            
            
            #=================
            
            #writer = pd.ExcelWriter("Calculation.xlsx", engine="xlsxwriter")
            ##worksheet = writer.add_worksheet()
            ##cell_format = writer.add_format({'bold': True, 'font_color': 'red'})
            ##cell_format.set_align('center')
            #pd_input.to_excel(writer, sheet_name=(f"{panel_name.upper()}"), startrow=0, header=True,startcol=0,float_format="%.2f")
            #fd_input.to_excel(writer, sheet_name=(f"{panel_name.upper()}"),startrow=7,header=True,startcol=0,float_format="%.2f")  # Default position, cell A1.
            #
            #writer.close()
            ##DataFrame.to_excel(excel_writer, *, sheet_name='Sheet1', na_rep='', float_format=None, columns=None, header=True, index=True, index_label=None, startrow=0, startcol=0, engine=None, merge_cells=True, inf_rep='inf', freeze_panes=None, storage_options=None, engine_kwargs=None)[source]
            
        def Add_row():

            f_type_var=tk.StringVar()
            f_power_var=tk.DoubleVar()
            f_phase_var = tk.StringVar()
            f_pf_var = tk.DoubleVar()
            f_cable_len_var=tk.DoubleVar()

            current_val=tk.DoubleVar()
            cable=tk.DoubleVar()
            f_phase_name_var=tk.StringVar()

            lbl_f_number=ttk.Label(panel_top5,text=len(rows)+1,width=4)
            lbl_f_number.grid(row=len(rows), column=0,padx=3,pady=2,sticky='NW')

            combo_f_type = ttk.Combobox(panel_top5, width=14,textvariable=f_type_var)
            combo_f_type.config(values=("Lighting","Socket","Equipment","Motor(1P-DOL)","Motor(3P-DOL)","Motor(3P-YD)","Panel"),state='readonly')
            combo_f_type.current(0)
            combo_f_type.grid(row=len(rows), column=1,padx=3,pady=2,sticky='NW')

            ent_f_power = ttk.Entry(panel_top5, width=12,textvariable=f_power_var)

            ent_f_power.grid(row=len(rows), column=2,padx=3,pady=2,sticky='NW')

            combo_f_phase = ttk.Combobox(panel_top5, width=8)
            combo_f_phase.config(textvariable=f_phase_var,values=[1, 3],state='readonly')
            combo_f_phase.current(0)
            combo_f_phase.grid(row=len(rows), column=3,padx=3,pady=2,sticky='NW')

            ent_f_pf = ttk.Entry(panel_top5, width=12,textvariable=f_pf_var)

            ent_f_pf.grid(row=len(rows), column=4,padx=3,pady=2,sticky='NW')

            ent_f_cable_len = ttk.Entry(panel_top5, width=12,textvariable=f_cable_len_var,validate='key')
            ent_f_cable_len.grid(row=len(rows), column=5,padx=3,pady=2,sticky='NW')


            lbl_f_current = ttk.Label(panel_top5,textvariable=current_val, width=12,background=cream,justify='center')
            lbl_f_current.grid(row=len(rows), column=6,padx=3,pady=2,sticky='NW')

            lbl_f_cable = ttk.Label(panel_top5,textvariable=cable, width=12,background=cream,justify='center')
            lbl_f_cable.grid(row=len(rows), column=7,padx=3,pady=2,sticky='NW')   

            lbl_f_phase_name = ttk.Label(panel_top5,textvariable=f_phase_name_var, width=12,background=cream,justify='center')
            lbl_f_phase_name.grid(row=len(rows), column=8,padx=3,pady=2,sticky='NW')

            lbl_f_delta_v = ttk.Label(panel_top5,text="", width=12,background=cream,justify='center')
            lbl_f_delta_v.grid(row=len(rows), column=9,padx=3,pady=2,sticky='NW') 

            lbl_f_breaker = ttk.Label(panel_top5,text="", width=12,background=cream,justify='center')
            lbl_f_breaker.grid(row=len(rows), column=10,padx=3,pady=2,sticky='NW') 
            lbl_f_name = ttk.Label(panel_top5,text="",width=12,background=cream)
            lbl_f_name.grid(row=len(rows), column=11,padx=3,pady=2,sticky='NW')
            lbl_f_bmetal = ttk.Label(panel_top5,text="",width=12,background=cream)
            lbl_f_bmetal.grid(row=len(rows), column=12,padx=3,pady=2,sticky='NW')
            lbl_f_setting = ttk.Label(panel_top5,text="",width=12,background=cream)
            lbl_f_setting.grid(row=len(rows), column=13,padx=3,pady=2,sticky='NW')
            lbl_f_contactor = ttk.Label(panel_top5,text="",width=12,background=cream)
            lbl_f_contactor.grid(row=len(rows), column=14,padx=3,pady=2,sticky='NW')



            rows.append([lbl_f_number,combo_f_type,ent_f_power,combo_f_phase,ent_f_pf,ent_f_cable_len,lbl_f_current,lbl_f_cable,lbl_f_phase_name,lbl_f_delta_v,lbl_f_breaker,lbl_f_name,lbl_f_bmetal,lbl_f_setting,lbl_f_contactor])

        def reset_feeders():
            for widget in panel_top5.winfo_children():
                    widget.destroy()
                    rows.clear()
        #==========================================Def Del Last Row==============================================            
        def del_row():
            last_row = rows.pop()
            for widget1 in last_row:
                widget1.destroy() 
            while  rows2:
                last_row2 = rows2.pop()   
                for widget in last_row2:
                    widget.destroy()
            #calculate()     
            #messagebox.showerror("NOTE","Please Press Again Calculate Bottun")
        #=========================================def panel save==========================================
        self.panel_d_f_var=tk.DoubleVar()
        self.main_cable_len_var=tk.IntVar()
        #self.temp_var=tk.IntVar()
        #self.instalation_var=tk.StringVar()
        #self.insulation_var=tk.StringVar()
        self.max_volage_drop_var=tk.DoubleVar()

        #========================================panel_top1(TOP Lables)======================================
        file_path = "data.xlsx"

        # بارگذاری دو شیت از اکسل
        sheets = {
        "In Air": pd.read_excel(file_path, sheet_name="In Air"),
        "In Ground": pd.read_excel(file_path, sheet_name="In Ground")
        }

        b=("Times", "16", "bold italic")
        lbl_input=ttk.Label(panel_top1,text="INPUTS",font=b,justify='left',background=blue_dark,foreground=white,width=8,anchor=E)

        lbl_panel_name=ttk.Label(panel_top1, text="Panel Name :",background='white',width=22)
        panel_name_input=ttk.Label(panel_top1,text=panel_name.upper(),background='white',width=8)

        lbl_panel_phase=ttk.Label(panel_top1, text="Panel phase :",background='white',width=22)
        panel_phase_input=ttk.Combobox(panel_top1,width=5)

        panel_phase_input.config(value=("RST","R","S","T"),state='readonly')
        panel_phase_input.current(0)

        lbl_upstream_panel=ttk.Label(panel_top1, text="Upstream Panel :",background='white',width=22)
        upstream_panel=ttk.Entry(panel_top1,background='white',width=8)

        lbl_panel_df=ttk.Label(panel_top1, text="Panel Demand Factor :",justify='left',background='white',width=22)
        panel_d_f_input=ttk.Entry(panel_top1,width=8,textvariable=self.panel_d_f_var ,background='white')
        lbl_maincable_len=ttk.Label(panel_top1, text="Main Cacle Lentgh(m) :",justify=LEFT,background='white',width=22)
        main_cable_len=ttk.Entry(panel_top1,width=8,background='white',textvariable=self.main_cable_len_var)
        lbl_temp=ttk.Label(panel_top1, text="Ambient Tempreture(C) :",justify='left',background='white',width=22)
        temp_intry=ttk.Combobox(panel_top1,width=5,justify=CENTER,values=list(range(10, 75, 5)))
        temp_intry.config(state='readonly')
        temp_intry.current(4)
        lbl_instalation=ttk.Label(panel_top1, text="Cable Installation :",background='white',width=22)
        instalation_input=ttk.Combobox(panel_top1,width=10, values=["In Air", "In Ground"])
        instalation_input.config(state='readonly')
        instalation_input.current(0)
        lbl_p_instalation=ttk.Label(panel_top1, text="Cable Insullation :",background='white',width=22)
        c_insulation_input=ttk.Combobox(panel_top1,width=10)
        c_insulation_input.config(state='readonly',values=["PVC", "XLPE"])
        c_insulation_input.current(0)
        lbl_max_volage_drop=ttk.Label(panel_top1, text="Max voltage drop:",background='white',width=22)
        max_volage_drop_input=ttk.Entry(panel_top1,width=8,text=self.max_volage_drop_var ,background=cream)

        #=========================================== Grids panel_top1============================================

        lbl_input.grid(row=0,column=0,padx=3,pady=3,rowspan=2,columnspan=1,sticky=W)
        lbl_panel_name.grid(row=0,column=1,rowspan=1,columnspan=1,padx=3,pady=3,sticky=W)
        panel_name_input.grid(row=0,column=2,rowspan=1,columnspan=1,padx=3,pady=3,sticky=W)
        space(panel_top1,0,3,blue_dark)
        lbl_panel_phase.grid(row=0,column=4,rowspan=1,columnspan=1,padx=3,pady=3,sticky=W)
        panel_phase_input.grid(row=0,column=5,rowspan=1,columnspan=1,padx=3,pady=3,sticky=W)
        lbl_panel_df.grid(row=0,column=7,rowspan=1,columnspan=1,padx=3,pady=3,sticky=W)
        panel_d_f_input.grid(row=0,column=8,rowspan=1,columnspan=1,padx=3,pady=3,sticky=W)
        lbl_upstream_panel.grid(row=1,column=1,rowspan=1,columnspan=1,padx=3,pady=3,sticky=W)
        space(panel_top1,0,6,blue_dark)

        upstream_panel.grid(row=1,column=2,rowspan=1,columnspan=1,padx=3,pady=3,sticky=W)
        lbl_maincable_len.grid(row=1,column=4,rowspan=1,columnspan=1,padx=3,pady=3,sticky=W)   
        main_cable_len.grid(row=1,column=5,rowspan=1,columnspan=1,padx=3,pady=3,sticky=W)
        main_cable_length=(main_cable_len.get())
        lbl_temp.grid(row=1,column=7,rowspan=1,columnspan=1,padx=3,pady=3,sticky=W)
        temp_intry.grid(row=1,column=8,rowspan=1,columnspan=2,padx=3,pady=3,sticky=W)
        space(panel_top1,0,9,blue_dark)
        lbl_instalation.grid(row=0,column=10,rowspan=1,columnspan=1,padx=3,pady=3,sticky=W)
        instalation_input.grid(row=0,column=11,rowspan=1,columnspan=2,padx=3,pady=3,sticky=W)

        lbl_p_instalation.grid(row=1,column=10,rowspan=1,columnspan=1,padx=3,pady=3,sticky=W)
        c_insulation_input.grid(row=1,column=11,rowspan=1,columnspan=2,padx=3,pady=3,sticky=W)
        space(panel_top1,0,13,blue_dark)
        lbl_max_volage_drop.grid(row=0,column=14,rowspan=1,columnspan=2,padx=3,pady=3,sticky=W)
        max_volage_drop_input.grid(row=0,column=16,rowspan=1,columnspan=2,padx=3,pady=3,sticky=W)
        
        #===========================================panel_top2(Resualt Panel)=================================================
        lbl_output=ttk.Label(panel_top2,text="OUTPUT",font=b,justify='left',background=blue_dark,foreground=white,width=8,anchor=E)
        lbl_total_power=ttk.Label(panel_top2,text="Conected Load(KW):",justify='left',background='white',width=23)
        lbl_panel_current=ttk.Label(panel_top2,text="Conected Current(A):",justify='left',background='white',width=23)
        lbl_total_d_power=ttk.Label(panel_top2,text="Demand Load(KW):",justify='left',background='white',width=22)
        lbl_panel_d_current=ttk.Label(panel_top2,text="Demand Current(A):",justify='left',background='white',width=22)
        lbl_total_PF=ttk.Label(panel_top2,text="Power Factor",justify='left',background='white',width=22)
        lbl_panel_cable=ttk.Label(panel_top2,text="Main Cable(mm2)",justify='left',background='white',width=22)
        lbl_panel_bearker=ttk.Label(panel_top2,text="Main Braker(A)",justify='left',background='white',width=22)
        lbl_panel_upstrream_cb=ttk.Label(panel_top2,text="Upstream Braker(A)",justify='left',background='white',width=22)
        lbl_panel_delta_v=ttk.Label(panel_top2,text="Delta V%",justify='left',background='white',width=22)

        #=========================================== Grids panel_top2============================================
        lbl_output.grid(row=0,column=0,padx=3,pady=3,rowspan=2,columnspan=1,sticky=W)

        lbl_total_power.grid(row=0,column=1,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)
        holder(panel_top2,0,2,8)
        space(panel_top2,0,3,green2)
        lbl_total_d_power.grid(row=0,column=4,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)
        holder(panel_top2,0,5,8)
        space(panel_top2,0,6,green2)
        lbl_total_PF.grid(row=0,column=7,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)
        holder(panel_top2,0,8,8)
        space(panel_top2,0,9,green2)
        lbl_panel_bearker.grid(row=0,column=10,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)
        holder(panel_top2,0,11,8)
        space(panel_top2,0,12,green2)
        space(panel_top2,0,13,green2)
        lbl_panel_upstrream_cb.grid(row=0,column=14,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)
        holder(panel_top2,0,15,8)
        space(panel_top2,0,16,green2)


        lbl_panel_current.grid(row=1,column=1,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)
        holder(panel_top2,1,2,8)
        space(panel_top2,1,3,green2)
        lbl_panel_d_current.grid(row=1,column=4,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)
        holder(panel_top2,1,5,8)  
        space(panel_top2,1,6,green2)
        lbl_panel_cable.grid(row=1,column=7,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)
        holder(panel_top2,1,8,8)
        space(panel_top2,1,9,green2)
        lbl_panel_delta_v.grid(row=1,column=10,padx=3,pady=3,rowspan=1,columnspan=1,sticky=W)
        holder(panel_top2,1,11,8)
        space(panel_top2,1,12,green2)

        #=========================================== LABLES panel_top4 ===================================
        lbl_num=ttk.Label(panel_top4, text="No",width=5,background=cream)
        lbl_feeder_type=ttk.Label(panel_top4, text="F.TYPE",justify='left',background=cream,width=17)
        lbl_power=ttk.Label(panel_top4, text="POWER(KW)",justify='left',background=cream,width=12)
        lbl_phase=ttk.Label(panel_top4, text="PHASE",justify='left',background=cream,width=12)
        lbl_power_factor=ttk.Label(panel_top4, text="P.FACTOR",justify='left',background=cream,width=12)
        lbl_lentgh=ttk.Label(panel_top4, text="LENTGH(m)",justify='left',background=cream,width=12)
        lbl_current=ttk.Label(panel_top4, text="CURRENT(A)",justify='left',background=cream,width=12)
        lbl_cable_size=ttk.Label(panel_top4, text="CABLE.SIZE",justify='left',background=cream,width=12)
        lbl_ph_name=ttk.Label(panel_top4, text="PH.NAME",justify='left',background=cream,width=12)
        lbl_delta_v=ttk.Label(panel_top4, text="DELTA V",justify='left',background=cream,width=12)
        lbl_feeder_breaker=ttk.Label(panel_top4, text="BREAKER",justify='left',background=cream,width=12)
        lbl_feeder_name=ttk.Label(panel_top4, text="F.NUMBER",justify='left',background=cream,width=12)
        lbl_bmetal=ttk.Label(panel_top4, text="BMetal(A)",justify='left',background=cream,width=12)
        lbl_setting=ttk.Label(panel_top4, text="Setting(A)",justify='left',background=cream,width=12)
        lbl_contactor=ttk.Label(panel_top4, text="Contactor(A)",justify='left',background=cream,width=12)

        #========================================= BOttom panel_top3============================================
        btn_new_feeder=Button(panel_top3,text="Add A Feeder",width=12, background=blue_light,command=lambda:Add_row())
        btn_calculate=Button(panel_top3,text="Calculate",width=20, background=green_fosfori,command=calculate)
        btn_del=Button(panel_top3, text="Del.Last Row",width=12,justify='left',background='yellow',command=del_row)
        btn_reset=Button(panel_top3, text="Reset Feeders",width=12,justify='left',background='red',command=reset_feeders)
        btn_change_name = tk.Button(panel_top3, text="Rename Panel",justify='left',width=12, command=lambda: change_tab_name())


        #========================================= GRID BOttom panel_top3============================================
        space(panel_top3,0,0,white)
        btn_new_feeder.grid(row=0,column=1,rowspan=1,columnspan=1,padx=3,pady=3,sticky='SNEW')
        space(panel_top3,0,2,white)
        btn_change_name.grid(row=0,column=3,padx=3,pady=3,sticky='SNEW')
        space(panel_top3,0,4,white)
        btn_del.grid(row=0,column=5,padx=3,pady=3,sticky='SNEW')
        space(panel_top3,0,6,white)
        btn_reset.grid(row=0,column=7,padx=3,pady=3,sticky='SNEW')
        space(panel_top3,0,8,white)
        btn_calculate.grid(row=0,column=9,rowspan=1,columnspan=1,padx=3,pady=3,sticky='SNEW')

        #======================================= Grid panel_top4 ===========================================
        lbl_num.grid(row=0,column=1,rowspan=1,columnspan=1,padx=3,pady=3,sticky='NW')
        lbl_feeder_type.grid(row=0,column=2,padx=3,pady=3,rowspan=1,columnspan=1,sticky='NW')
        lbl_power.grid(row=0,column=3,padx=3,pady=3,rowspan=1,columnspan=1,sticky='NW')
        lbl_phase.grid(row=0,column=4,padx=3,pady=3,rowspan=1,columnspan=1,sticky='NW')
        lbl_power_factor.grid(row=0,column=5,padx=3,pady=3,rowspan=1,columnspan=1,sticky='NW')
        lbl_lentgh.grid(row=0,column=6,padx=3,pady=3,rowspan=1,columnspan=1,sticky='NW')
        lbl_current.grid(row=0,column=7,padx=3,pady=3,rowspan=1,columnspan=1,sticky='NW')
        lbl_cable_size.grid(row=0,column=8,padx=3,pady=3,rowspan=1,columnspan=1,sticky='NW')
        lbl_ph_name.grid(row=0,column=9,padx=3,pady=3,rowspan=1,columnspan=1,sticky='NW')
        lbl_delta_v.grid(row=0,column=10,padx=3,pady=3,rowspan=1,columnspan=1,sticky='NW')
        lbl_feeder_breaker.grid(row=0,column=11,padx=3,pady=3,rowspan=1,columnspan=1,sticky='NW')
        lbl_feeder_name.grid(row=0,column=12,padx=3,pady=3,rowspan=1,columnspan=1,sticky='NW')
        lbl_bmetal.grid(row=0,column=13,padx=3,pady=3,rowspan=1,columnspan=1,sticky='NW')
        lbl_setting.grid(row=0,column=14,padx=3,pady=3,rowspan=1,columnspan=1,sticky='NW')
        lbl_contactor.grid(row=0,column=15,padx=3,pady=3,rowspan=1,columnspan=1,sticky='NW')

        def change_tab_name():
            active_tab_id = self.notebook.select()
            new_name_entry = simpledialog.askstring("Rename Panel", "Enter the new name for Panel")
            self.notebook.tab(active_tab_id, text=(new_name_entry.upper()))
            panel_name_input.config(text=(new_name_entry.upper()))
    
    def delete_active_tab(self):
        confirm_delete = messagebox.askyesno("Delete Tab", "Are you sure you want to delete this Panel?")
        if confirm_delete:
            active_tab_id = self.notebook.select()
            self.notebook.forget(active_tab_id)


    
    
   

    
    
    


    

mroot.mainloop()