import tkinter as tk
from tkinter import ttk, simpledialog, messagebox

##############################
from openpyxl import *
from openpyxl import Workbook
from xlsxwriter import Workbook
from tkinter import messagebox
from tkinter import SCROLL
from tkinter import ttk, simpledialog
######################################

COLORS = {
    "cream": '#dad7cd',
    "white": 'white',
    "blue_dark": '#00264d',
    "blue_light": '#b3f0ff',
    "green_fosfori": '#4dff4d',
    "green2": '#a3b18a',
    "green4": '#3a5a40',
    "green5": '#344e41'
}
#[[[[[[[[[[[[[[[[[[[[[[[[[[[
project_list=[]
#]]]]]]]]]]]]]]]]]]]]]]]]]]]
class PanelProject:
    """Class to manage panel projects."""

    def __init__(self, root):
        self.root = root
        self.create_window()
    #{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{{
        
    

   
        
    def bar(self):
        """Creates the main menu bar."""
        menubar = tk.Menu(self.root)

        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="New", command=lambda: None)
        filemenu.add_command(label="Open", command=lambda: None)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=filemenu)

        editmenu = tk.Menu(menubar, tearoff=0)
        editmenu.add_command(label="Change Language", command=self.change_language)
        menubar.add_cascade(label="Edit", menu=editmenu)

        self.root.config(menu=menubar)   
        
        
           
    #}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}

    def create_window(self):
        """Creates the Panel Project window."""
        self.window = tk.Toplevel(self.root)
        #self.window.title("Panel Project")
        #self.window.geometry("600x400")
        #self.window.configure(bg=COLORS["cream"])
#
        ## Title
        #title_label = ttk.Label(
        #    self.window, text="Panel Project Management",
        #    font=("Corbel", 16, "bold"), background=COLORS["cream"]
        #)
        #title_label.pack(pady=10)
#
        ## Add Button
        #add_button = tk.Button(
        #    self.window, text="Add Panel", bg=COLORS["green_fosfori"],
        #    command=self.add_panel
        #)
        #add_button.pack(pady=5)
#
        ## Delete Button
        #delete_button = tk.Button(
        #    self.window, text="Delete Panel", bg=COLORS["blue_light"],
        #    command=self.delete_panel
        #)
        #delete_button.pack(pady=5)
        
##################################################################       
        project_name = simpledialog.askstring("New Tab", "Enter the name for the Project")
        while  project_name=='':
            ValueError
            messagebox.showerror("Error", "Please Insert a Name for Project")
            project_name = simpledialog.askstring("New Tab", "Enter the name for the Project")   
              
        project_name=project_name
        
        self.window.title(f"{project_name} Project")
        self.window.geometry("1300x850+100+50")
        self.window.state("zoomed") 
       
        
        self.header=ttk.Frame(self.window,bg=COLORS['blue_dark'],height=200,width=self.window.winfo_screenwidth())
        self.header.grid(row=0,column=0,rowspan=1,columnspan=1,sticky='snew')
        
        self.under_header=ttk.Frame(self.window,bg=COLORS['green2'],height=2,width=self.window.winfo_screenwidth())
        self.under_header.grid(row=1,column=0,rowspan=1,columnspan=1,sticky='snew')
        
        self.notebook = ttk.Notebook(self.window,height=620,width=self.window.winfo_screenwidth())# ایجاد نوت‌بوک برای تب‌ها
        self.notebook.grid(row=3,column=0,rowspan=1,columnspan=1,sticky='snew') 
                 
        self.tab_control = ttk.Notebook(self.window,height=620,width=self.window.winfo_screenwidth())
        self.panel_list=[] # لیست برای ذخیره نام تب‌ها
        
        
        a=("Helvetica", "10", "bold")
        c=("Literal", "12", "bold")
    
        self.space_header=tk.Label(self.header,text="",width=26,background=COLORS['blue_dark'])
        self.add_button = tk.Button(self.header, text="Add Panel",width=24,font=a,background=COLORS['green4'],foreground=COLORS['white'],command=self.add_panel)
        self.delete_button = tk.Button(self.header, text="Delete Panel",width=24,justify='left',font=a,background=COLORS['green4'],foreground=COLORS['white'],command=self.delete_active_tab)
        self.excell_button = tk.Button(self.header, text="To Excell",width=24,justify='left',font=a,background=COLORS['green4'],foreground=COLORS['white'],command=self.to_excell)
        self.pdf_button=tk.Button(self.header, text="To PDF",width=24,justify='left',font=a,background=COLORS['green4'],foreground=COLORS['white'])#,command=self.to_PDF)
        self.company_lbl=tk.Label(self.header,text="گروه نرم افزاری مانی نیروی البرز",justify='left', font=('Vazir','16','bold'),foreground="white",background=COLORS['blue_dark'])
        
        self.add_button.grid(row=0,column=1,padx=5,pady=5,sticky='snew')
        self.delete_button.grid(row=0,column=2,padx=5,pady=5,sticky='snew')
        self.excell_button.grid(row=0,column=3,padx=5,pady=5,sticky='snew')
        self.pdf_button.grid(row=0,column=4,padx=5,pady=5,sticky='snew')
        self.space_header.grid(row=0,column=5,padx=5,pady=5,sticky='snew')
        self.company_lbl.grid(row=0,column=6,padx=5,pady=5,sticky='snew')
        self.bar()
        project_list.append(project_name)
        print(project_list)
        
        
        ptab = ttk.Frame(self.notebook,width=self.root.winfo_screenwidth(),height=600)
        self.notebook.add(ptab, text='Project Info('+project_name+')')
        self.notebook.select(ptab)
        

        project_top0=ttk.Frame(ptab,bg=COLORS['blue_dark'],border=2,height=100,width=1500) # پنجره panel_top0
        project_top0.grid_propagate(False)
        project_top0.grid(row=0,column=0,rowspan=1,columnspan=1,sticky='snew')
        
        
        project_top1=ttk.Frame(ptab,bg=COLORS['blue_dark'],border=2,height=550,width=1500)           # پنجره panel_top1
        project_top1.grid_propagate(False)
        project_top1.grid(row=1,column=0,rowspan=1,columnspan=1,sticky='sn')

        
       
        
        
        
        lbl_project_prop=ttk.Label(project_top0, text="Project Information:",font=('Literal','22'),justify='center',background=COLORS['blue_dark'],foreground=COLORS['white'],width=22)
        
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
        
        
        lbl_space=ttk.Label(project_top1, text="",justify='center',background=COLORS['blue_dark'],width=50)
        lbl_space.grid(row=1,column=0,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')
        
        lbl_space2=ttk.Label(project_top1, text="",justify='center',background=COLORS['blue_dark'],width=50)
        lbl_space2.grid(row=1,column=6,rowspan=1,columnspan=1,padx=10,pady=10,sticky='snew')
        
        lbl_space3=ttk.Label(project_top1, text="",justify='center',background=COLORS['blue_dark'],width=50)
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
        
         
        
        
#####################################################################
    #def add_panel(self):
    #    """Adds a new panel to the project."""
    #    panel_name = simpledialog.askstring("New Panel", "Enter the name of the panel:")
    #    if panel_name:
    #        messagebox.showinfo("Panel Added", f"Panel '{panel_name}' has been added.")
            
            
    def add_panel(self):
        panel_name = simpledialog.askstring("New Tab", "Enter the name for the new Panel")
        tab = ttk.Frame(self.notebook,width=self.root.winfo_screenwidth(),height=600)
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
        
        panel_top1=ttk.Frame(tab,bg=COLORS["blue_dark"],border=2)           # پنجره panel_top1
        panel_top1.grid(row=0,column=0,rowspan=1,columnspan=1,sticky='SNEW')

        panel_top2=ttk.Frame(tab,bg=COLORS["blue_dark"],border=2)           # پنجره panel_top2
        panel_top2.grid(row=1,column=0,rowspan=1,columnspan=1,sticky='SNEW')

        panel_top3=ttk.Frame(tab,bg=COLORS["white"],border=2)           # پنجره panel_top3
        panel_top3.grid(row=2,column=0,rowspan=1,columnspan=1,sticky='SNEW')

        panel_top4=ttk.Frame(tab,bg=COLORS['cream'],border=2)            # پنجره panel_top4
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
      
            
            
            

    def delete_panel(self):
        """Deletes a panel from the project."""
        panel_name = simpledialog.askstring("Delete Panel", "Enter the name of the panel to delete:")
        if panel_name:
            messagebox.showinfo("Panel Deleted", f"Panel '{panel_name}' has been deleted.")