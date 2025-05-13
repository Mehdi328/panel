import tkinter as tk
from tkinter import ttk, simpledialog, messagebox, filedialog

from logic import PanelCalculator,ReportGenerator
from ups_battery_ampere import UPSBatteryAmpere
from ups_support_time import UPSBatterySupport

import re
from utilities import COLORS

# ========================================== Colors =====================================================
cream = '#dad7cd'
white = 'white'
blue_dark = '#00264d'
blue_light = '#b3f0ff'
green_fosfori = '#4dff4d'
green2 = '#a3b18a'
green4 = '#3a5a40'
green5 = '#344e41'
project_list=[]


class Feeder:
    def __init__(self, parent, row_idx, panel_calculator):
        self.parent = parent
        self.row_idx = row_idx
        self.panel_calculator = panel_calculator
        
        self.f_type_var = tk.StringVar()
        self.f_power_var = tk.DoubleVar()
        self.f_phase_var = tk.StringVar()
        self.f_pf_var = tk.DoubleVar()
        self.f_cable_len_var = tk.DoubleVar()
        
        self.f_current = tk.DoubleVar()
        self.f_cable = tk.DoubleVar()
        self.f_phase_name = tk.StringVar()
        self.f_deltav = tk.DoubleVar()
        self.f_breaker = tk.DoubleVar()
        self.f_bmetal = tk.DoubleVar()
        self.f_setting = tk.DoubleVar()
        self.f_contactor = tk.DoubleVar()
        self.f_name = tk.StringVar()
        
        self.widgets = []
        self.calculate_callback = self.get_data  # تغییر به متد محلی
        self.create_widgets()

    # [متدهای مربوط به رابط کاربری Feeder]
    def create_widgets(self):
        """ایجاد ویجت‌های فیدر."""
        lbl_number = ttk.Label(self.parent, text=str(self.row_idx + 1), width=3, justify='center')
        combo_f_type = ttk.Combobox(self.parent, width=14, textvariable=self.f_type_var, values=["Lighting", "Socket", "Equipment", "Motor(1P-DOL)", "Motor(3P-DOL)", "Motor(3P-YD)", "Panel"], state='readonly')
        combo_f_type.current(0)
        ent_power = ttk.Entry(self.parent, width=12, textvariable=self.f_power_var)
        combo_phase = ttk.Combobox(self.parent, width=6, textvariable=self.f_phase_var, values=[1, 3], state='readonly')
        combo_phase.current(0)
        ent_pf = ttk.Entry(self.parent, width=10, textvariable=self.f_pf_var)
        ent_cable_len = ttk.Entry(self.parent, width=10, textvariable=self.f_cable_len_var)
        
        lbl_f_current=tk.Label(self.parent,text="---",textvariable=self.f_current,width=10, background=cream)
        lbl_f_cable=tk.Label(self.parent,text="---",textvariable=self.f_cable,width=10, background=cream)
        lbl_f_phase_name=tk.Label(self.parent,text="---",textvariable=self.f_phase_name,width=10, background=cream)
        lbl_f_delta_v=tk.Label(self.parent,text="---",textvariable=self.f_deltav,width=10, background=cream)
        lbl_f_breaker=tk.Label(self.parent,text="---",textvariable=self.f_breaker,width=10, background=cream)
        lbl_f_bmetal=tk.Label(self.parent,text="---",textvariable=self.f_bmetal,width=10, background=cream)
        lbl_f_setting=tk.Label(self.parent,text="---",textvariable=self.f_setting,width=10, background=cream)
        lbl_f_contactor=tk.Label(self.parent,text="---",textvariable=self.f_contactor,width=12, background=cream)
        lbl_f_name=tk.Label(self.parent,text="---",textvariable=self.f_name,width=12, background=cream)
      

         #اتصال رویدادها برای محاسبات لحظه‌ای
        #ent_power.bind("<KeyRelease>", lambda e: self.calculate_callback(self.get_data()))
        #combo_phase.bind("<<ComboboxSelected>>", lambda e: self.calculate_callback(self.get_data()))
        #ent_pf.bind("<KeyRelease>", lambda e: self.calculate_callback(self.get_data()))
        #ent_cable_len.bind("<KeyRelease>", lambda e: self.calculate_callback(self.get_data()));
        for widget in [ent_power, ent_pf, ent_cable_len]:
            widget.bind("<KeyRelease>", lambda e: self.trigger_panel_calculation())
        combo_phase.bind("<<ComboboxSelected>>", lambda e: self.trigger_panel_calculation())


        # ذخیره ویجت‌ها
        self.widgets = [lbl_number, combo_f_type, ent_power, combo_phase, ent_pf, ent_cable_len,lbl_f_current,lbl_f_cable,lbl_f_phase_name,lbl_f_delta_v,lbl_f_breaker,lbl_f_bmetal,lbl_f_setting,lbl_f_contactor,lbl_f_name]

        # قرار دادن ویجت‌ها در گرید
        for col, widget in enumerate(self.widgets):
            widget.grid(row=self.row_idx+1, column=col, padx=2, pady=3)
    
    def calculate_callback(self, feeder_data):
        """فراخوانی متد محاسباتی با داده‌های فیدر."""
        try:
            self.panel_calculator.calculate_feeders(feeder_data)

        except Exception as e:
            messagebox.showerror("Error", f"Calculation failed: {e}")
        
    def trigger_panel_calculation(self):
        """فراخوانی محاسبات مربوط به فیدرها."""
        #try:
        if hasattr(self.panel_calculator, 'calculate_feeders'):
            self.panel_calculator.calculate_feeders()
        else:
            raise AttributeError("PanelCalculator does not have a method named 'calculate_feeders'.")
        #except Exception as e:
        #    messagebox.showerror("Error", f"Calculation failed: {e}")
    def f_out_widget(self):
        #current=logic.calculate_feeders(["current"])
        current="f.current"
        cable="f.cable"
        f_phase_name="f.phase name"
        breaker="f.Breaker"
        delta_v="f_.elta v"
        bmetal="f.bmetal"
        setting="f.setting"
        contactor="f.contactor"   
        lbl_f_current_c = ttk.Label(self.panel_top5,text=#"{:.2f}".format
                                    (current), width=12,background=green_fosfori,justify='center')
        lbl_f_current_c.grid(row=len(self.rows2), column=6,padx=3,pady=2,sticky='NW')
        lbl_f_cable_c = ttk.Label(self.panel_top5,text=cable, width=12,
                                  background=green_fosfori,justify='center')
        lbl_f_cable_c.grid(row=len(self.rows2), column=7,padx=3,pady=2,sticky='NW')  
        lbl_f_phase_name_c = ttk.Label(self.panel_top5,text=f_phase_name, width=12,
                                       background=green_fosfori,justify='center')
        lbl_f_phase_name_c.grid(row=len(self.rows2), column=8,padx=2,pady=3,sticky='NW')
        lbl_f_delta_v_c = ttk.Label(self.panel_top5,text=#"{:.2f}".format
                                    (delta_v), width=12,background=green_fosfori,justify='center')
        lbl_f_delta_v_c.grid(row=len(self.rows2), column=9,padx=3,pady=2,sticky='NW')
        lbl_f_breaker_c = ttk.Label(self.panel_top5,text=breaker, width=12,
                                    background=green_fosfori,justify='center')
        lbl_f_breaker_c.grid(row=len(self.rows2), column=10,padx=3,pady=2,sticky='NW')  
        lbl_f_name_c = ttk.Label(self.panel_top5,text=(f"F{len(self.rows2)+1}"),width=12,
                                 background=green_fosfori)
        lbl_f_name_c.grid(row=len(self.rows2), column=11,padx=3,pady=2,sticky='NW')
        lbl_f_bmetal_c=ttk.Label(self.panel_top5,text=bmetal,width=12,
                                 background=green_fosfori)
        lbl_f_bmetal_c.grid(row=len(self.rows2), column=12,padx=3,pady=2,sticky='NW')
        lbl_f_setting_c=ttk.Label(self.panel_top5,text=setting,width=12,
                                  background=green_fosfori)
        lbl_f_setting_c.grid(row=len(self.rows2), column=13,padx=3,pady=2,sticky='NW')
        lbl_f_contactor_c=ttk.Label(self.panel_top5,text=contactor,width=12,
                                    background=green_fosfori)
        lbl_f_contactor_c.grid(row=len(self.rows2), column=14,padx=3,pady=2,sticky='NW')
        
        self.rows2.append([lbl_f_current_c,lbl_f_cable_c,lbl_f_phase_name_c,lbl_f_delta_v_c,
                           lbl_f_breaker_c,lbl_f_name_c,lbl_f_bmetal_c,lbl_f_setting_c,
                           lbl_f_contactor_c])
            
    
    
    
    def get_data(self):
        """بازگرداندن داده‌های فیدر."""
        return {
            "type": self.f_type_var.get(),
            "power": float(self.f_power_var.get()) if self.f_power_var.get() else 0.0,
            "phase": int(self.f_phase_var.get()) if self.f_phase_var.get() else 1,
            "pf": float(self.f_pf_var.get()) if self.f_pf_var.get() else 0.0,
            "cable_len": float(self.f_cable_len_var.get()) if self.f_cable_len_var.get() else 0.0
        }
    
        #########################

    def validate_power(entry_widget):
        """اعتبارسنجی مقدار power."""
        try:
            value = float(entry_widget.get())
            if value > 0:
                entry_widget.configure(style="Valid.TEntry")  # مقدار معتبر
            else:
                entry_widget.configure(style="Invalid.TEntry")  # مقدار نامعتبر
        except ValueError:
            entry_widget.configure(style="Invalid.TEntry")  # مقدار نامعتبر (اگر عدد نباشد)
            
    def validate_cable_length(entry_widget):
        """اعتبارسنجی مقدار cable length."""
        try:
            value = float(entry_widget.get())
            if value > 0:
                entry_widget.configure(style="Valid.TEntry")  # مقدار معتبر
            else:
                entry_widget.configure(style="Invalid.TEntry")  # مقدار نامعتبر
        except ValueError:
            entry_widget.configure(style="Invalid.TEntry")  # مقدار نامعتبر (اگر عدد نباشد)
            
    def validate_power_factor(entry_widget):
        """اعتبارسنجی مقدار power factor."""
        try:
            value = float(entry_widget.get())
            if 0 < value <= 1:
                entry_widget.configure(style="Valid.TEntry")  # مقدار معتبر
            else:
                entry_widget.configure(style="Invalid.TEntry")  # مقدار نامعتبر
        except ValueError:
            entry_widget.configure(style="Invalid.TEntry")  # مقدار نامعتبر (اگر عدد نباشد)
            
    def validate_phase(panel_phase_widget, phase_widget):
        """اعتبارسنجی مقدار phase."""
        panel_phase = panel_phase_widget.get()
        phase = phase_widget.get()
        if panel_phase in ["R", "S", "T"] and phase == "3":
            phase_widget.configure(style="Invalid.TCombobox")  # مقدار نامعتبر
        else:
            phase_widget.configure(style="Valid.TCombobox")  # مقدار معتبر
            
    def validate_demand_factor(entry_widget):
        """اعتبارسنجی مقدار demand factor."""
        try:
            value = float(entry_widget.get())
            if 0 < value <= 1:
                entry_widget.config(bg="white")  # مقدار معتبر
            else:
                entry_widget.config(bg="red")  # مقدار نامعتبر
        except ValueError:
            entry_widget.config(bg="red")  # مقدار نامعتبر (اگر عدد نباشد)
        
    def validate_positive_number(entry_widget):
        """اعتبارسنجی مقدار power."""
        try:
            value = float(entry_widget.get())
            if value > 0:
                entry_widget.config(bg="white")  # مقدار معتبر
            else:
                entry_widget.config(bg="red")  # مقدار نامعتبر
        except ValueError:
            entry_widget.config(bg="red")  # مقدار نامعتبر (اگر عدد نباشد)
        
        
        
        #############################
    def destroy(self):
        """حذف ویجت‌های فیدر."""
        for widget in self.widgets:
            widget.destroy()

class PanelTab(ttk.Frame):
    def __init__(self, parent, panel_name, project_info):
        super().__init__(parent)
        self.panel_name = panel_name
        self.project_info = project_info
        self.parent = parent
        self.rows2 = []
        self.outputs = []
        self.feeders = []
        
        
        # ابتدا فیلدهای ورودی را تعریف می‌کنیم
        self.panel_phase_input = ttk.Combobox(self, values=["RST", "R", "S", "T"])
        self.panel_d_f_input = ttk.Entry(self)
        self.main_cable_len_input = ttk.Entry(self)
        self.temp_input = ttk.Entry(self)
        self.installation_input = ttk.Combobox(self, values=["In Air", "In Ground"])
        self.insulation_input = ttk.Combobox(self, values=["PVC", "XLPE"])
        self.max_vdrop_input = ttk.Entry(self)
        
        panel_data_values = {
            "panel_phase": self.panel_phase_input.get(),
            "demand_factor": float(self.panel_d_f_input.get() or 0.8),  # مقدار پیش‌فرض 0.8
            "main_cable_len": float(self.main_cable_len_input.get() or 10),
            "temp": int(self.temp_input.get() or 30),
            "installation": self.installation_input.get(),
            "insulation": self.insulation_input.get(),
            "max_vdrop": float(self.max_vdrop_input.get() or 5),
        }

        self.panel_calculator = PanelCalculator(**panel_data_values, project_info=self)
        self.panel_calculator.feeders = self.feeders
        
        
        self.create_inputs_section()
        self.create_output_section()
        self.create_buttons_section()
        self.create_feeders_table()
        self.create_feeder_header()
        self.initialize_settings()
        
        #self.panel_calculator = PanelCalculator(self.panel_phase_input,
        #    self.panel_d_f_input,
        #    self.main_cable_len_input,
        #    self.temp_input,
        #    self.installation_input,
        #    self.insulation_input,
        #    self.max_vdrop_input,
        #    project_info=self)  # ایجاد نمونه از PanelCalculator
#
        #self.panel_calculator.feeders = self.feeders

        
        

            # [تمام متدهای مربوط به رابط کاربری PanelTab اینجا قرار میگیرد]
    # (create_inputs_section, create_output_section, create_buttons_section و...)
    def initialize_settings(self):
        """مقداردهی اولیه متغیرها"""
        self.feeders.clear()
        self.rows2.clear()
        self.feeders_data = []
        self.current_list = [0, 0, 0]
        self.total_power = 0
        self.demand_current = 0

    def create_feeders_table(self):
        """ساخت جدول فیدرها با اسکرول"""
        self.scroll_frame = tk.Frame(self)
        self.scroll_frame.grid(row=4, column=0, sticky='nsew')

        self.canvas = tk.Canvas(self.scroll_frame)
        self.scrollbar = ttk.Scrollbar(self.scroll_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set,height=420)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.panel_top5 = self.scrollable_frame

    def create_inputs_section(self):

        """ساخت بخش ورودی اطلاعات تابلو"""
        self.input_frame = tk.Frame(self, bg=blue_dark, width=self.winfo_screenwidth(), height=65)
        self.input_frame.grid(row=0, column=0, sticky='SNW')
        self.input_frame.grid_propagate(False)

        font_label = ('Literal', 10, 'bold')
        entry_bg = 'white'

        # فیلدهای ورودی
        lbl = tk.Label(self.input_frame, text="INPUTS ", bg=blue_dark, justify='left', fg='white',width=7, font=('Helvetica', 16, 'bold'))
        lbl.grid(row=0, column=0, padx=2, pady=5, sticky='SNW', rowspan=2)
        labels = [
            ("Panel Name:", 0, 1),
            ("Panel Phase:", 0, 3),
            ("Demand Factor:", 0, 5),
            ("Cable Installation:", 0, 7),
            ("Max Voltage Drop(%):", 0, 9),
            ("Upstream Panel:", 1, 1),
            ("Main Cable Length(m):", 1, 3),
            ("Ambient Temp(c):", 1, 5),
            ("Cable Insulation:", 1, 7),
            
        ]

        for text, r, c in labels:
            lbl = tk.Label(self.input_frame, text=text,width=22, bg=blue_dark, fg='white', justify='left', font=font_label)
            lbl.grid(row=r, column=c, padx=2, pady=2, sticky='SNW')

        # ویجت‌های ورودی
        self.panel_name_input = tk.Entry(self.input_frame, text=self.panel_name.upper(), bg=entry_bg, width=8)
        self.panel_name_input.grid(row=0, column=2, padx=2, pady=5, sticky='SNW')
        
        self.panel_phase_input = ttk.Combobox(self.input_frame, values=["RST", "R", "S", "T"], width=5, state='readonly')
        self.panel_phase_input.current(0)
        self.panel_phase_input.grid(row=0, column=4, padx=2, pady=5, sticky='SNW')
    

        self.panel_d_f_input = tk.Entry(self.input_frame, width=8, bg=entry_bg)
        self.panel_d_f_input.insert(0, "0.75")  # مقدار دیفالت
        self.panel_d_f_input.grid(row=0, column=6, padx=2, pady=5, sticky='SNW')
        
        
        self.installation_input = ttk.Combobox(self.input_frame, values=["In Air", "In Ground"], width=10, state='readonly')
        self.installation_input.current(0)
        self.installation_input.grid(row=0, column=8, padx=2, pady=5, sticky='SNW')
        
        self.max_vdrop_input = tk.Entry(self.input_frame, width=8, bg=entry_bg)
        self.max_vdrop_input.insert(0, "5")  # مقدار دیفالت
        self.max_vdrop_input.grid(row=0, column=10, padx=2, pady=5, sticky='SNW')
        
        self.upstream_name_var = tk.StringVar()
        self.upstream_panel_input = tk.Entry(self.input_frame,textvariable=self.upstream_name_var, width=8, bg=entry_bg)
        self.upstream_panel_input.grid(row=1, column=2, padx=2, pady=5, sticky='SNW')
        

        self.main_cable_len_input = tk.Entry(self.input_frame, width=8, bg=entry_bg)
        self.main_cable_len_input.insert(0, "10")  # مقدار دیفالت
        self.main_cable_len_input.grid(row=1, column=4, padx=2, pady=5, sticky='SNW')
        

        self.temp_input = ttk.Combobox(self.input_frame, values=list(range(10, 75, 5)), width=5, state='readonly')
        self.temp_input.current(4)
        self.temp_input.grid(row=1, column=6, padx=2, pady=5, sticky='SNW')

        self.insulation_input = ttk.Combobox(self.input_frame, values=["PVC", "XLPE"], width=10, state='readonly')
        self.insulation_input.current(0)
        self.insulation_input.grid(row=1, column=8, padx=2, pady=5, sticky='SNW')

    def create_output_section(self):
        """ساخت بخش خروجی اطلاعات تابلو"""
        self.output_frame = tk.Frame(self, bg=green4, width=self.winfo_screenwidth(), height=65)
        self.output_frame.grid(row=1, column=0, sticky='SNW')
        self.output_frame.grid_propagate(False)
#
        font_label = ('Literal', 10, 'bold')
        entry_bg = 'white'

        # فیلدهای خروجی
        lbl = tk.Label(self.output_frame, text="OUTPUTS", bg=green4, fg='white',width=7, font=('Helvetica', 16, 'bold'))
        lbl.grid(row=0, column=0, padx=2, pady=5, sticky='SNW', rowspan=2)
        labels = [
            ("Conected Load(KW):", 0, 1),
            ("Conected Current(A):", 0, 3),
            ("Power Factor:", 0, 5),
            ("Main Breaker(A):", 0, 7),
            ("Delta V (%):", 0, 9),
            ("Demand Load(KW):", 1, 1),
            ("Demand Current(A):", 1, 3),
            ("Main Cable Size (mm²):", 1, 5),
            ("Upstream Breaker(A):", 1, 7),

            
        ]

        for text, r, c in labels:
            lbl = tk.Label(self.output_frame, text=text,width=22, bg=green4, fg='white', justify='left', font=font_label)
            lbl.grid(row=r, column=c, padx=2, pady=2, sticky='SNW')
            
        # ویجت‌های خروجی
        self.out_total_power = tk.Label(self.output_frame, text="", width=6, bg='white', justify='left')
        self.out_total_power.grid(row=0, column=2, padx=2, pady=5, sticky='SNW')

        self.out_panel_current = tk.Label(self.output_frame, text="", width=7, bg='white', justify='left')
        self.out_panel_current.grid(row=0, column=4, padx=2, pady=5, sticky='SNW')

        self.out_total_PF = tk.Label(self.output_frame, text="", width=7, bg='white', justify='left')
        self.out_total_PF.grid(row=0, column=6, padx=2, pady=5, sticky='SNW')

        self.out_panel_breaker = tk.Label(self.output_frame, text="", width=10, bg='white', justify='left')
        self.out_panel_breaker.grid(row=0, column=8, padx=2, pady=5, sticky='SNW')

        self.out_panel_delta_v = tk.Label(self.output_frame, text="", width=6, bg='white', justify='left')
        self.out_panel_delta_v.grid(row=0, column=10, padx=2, pady=5, sticky='SNW')

        self.out_total_d_power = tk.Label(self.output_frame, text="", width=6, bg='white', justify='left')
        self.out_total_d_power.grid(row=1, column=2, padx=2, pady=5, sticky='SNW')

        self.out_panel_d_current = tk.Label(self.output_frame, text="", width=7, bg='white', justify='left')
        self.out_panel_d_current.grid(row=1, column=4, padx=2, pady=5, sticky='SNW')

        self.out_panel_cable = tk.Label(self.output_frame, text="", width=7, bg='white', justify='left')
        self.out_panel_cable.grid(row=1, column=6, padx=2, pady=5, sticky='SNW')

        self.out_panel_upstream_cb = tk.Label(self.output_frame, text="", width=10, bg='white', justify='left')
        self.out_panel_upstream_cb.grid(row=1, column=8, padx=2, pady=5, sticky='SNW')

    def create_buttons_section(self):
        """ساخت دکمه‌های کنترلی"""
        self.buttons_frame = tk.Frame(self, bg='white',width=self.winfo_screenwidth(), height=50)
        self.buttons_frame.grid(row=2, column=0, rowspan=1,columnspan=1,padx=3,pady=3,sticky='SNEW')
        self.buttons_frame.grid_propagate(False)

        self.btn_add_row = tk.Button(self.buttons_frame, text="Add A Feeder", command=self.add_feeder, width=15, bg='#b3f0ff')
        self.btn_add_row.pack(side='left', padx=5, pady=5)

        self.btn_calculate = tk.Button(self.buttons_frame, text="Calculate", command=self.panel_calculator.calculate, width=15, bg='#4dff4d')
        self.btn_calculate.pack(side='left', padx=5, pady=5)

        self.btn_del_row = tk.Button(self.buttons_frame, text="Delete Last Row", command=self.del_row, width=15, bg='yellow')
        self.btn_del_row.pack(side='left', padx=5, pady=5)

        self.btn_reset = tk.Button(self.buttons_frame, text="Reset Feeders", command=self.reset_feeders, width=15, bg='red')
        self.btn_reset.pack(side='left', padx=5, pady=5)
        
    """نمایش خروجی فیدرها به صورت لیبل در زیر ستون‌های مربوطه."""
    def create_feeder_header(self):
    #    """ساخت بخش هدر برای فیدر هاو"""
    #    self.f_header_frame = tk.Frame(self, bg=cream,height=25)
    #    self.f_header_frame.grid(row=3, column=0, rowspan=1, columnspan=1, sticky='SNEW')
    #    self.f_header_frame.grid_propagate(False)
#
    #    #font_label = ('Literal', 10, 'bold')
    #    #entry_bg = 'white'
    #    
#
    #    
    #    labels = [
    #        ("No", 0, 0,3),
    #        ("F.TYPE", 0, 1,14),
    #        ("POWER(KW)", 0, 2,10),
    #        ("PHASE", 0, 3,10),
    #        ("P.FACTOR", 0, 4,10),
    #        ("LENGTH(m)", 0, 5,10),
    #        ("CURRENT(A)", 0, 6,10),
    #        ("CABLE.SIZE", 0, 7,10),
    #        ("PH.NAME", 0, 8,10),
    #        ("DELTA V", 0, 9,10),
    #        ("BREAKER", 0, 10,10),
    #        ("BMetal(A)", 0, 11,10),
    #        ("SETTING(A)", 0, 12,10),
    #        ("CONTACTOR(A)", 0, 13,12), 
    #        ("FEEDER NAME", 0, 14,10),
    #    ]
#
    #    for text, r, c, w in labels:
    #        
    #        
    #        
    #        
    #        
    #        
    #        lbl = tk.Label(self.f_header_frame, text=text,width=w, bg=cream, justify='center')
    #        lbl.grid(row=r, column=c,padx=2,pady=3,rowspan=1,columnspan=1,sticky='NW')
            
    ################################33
    
    
        #"""ساخت بخش هدر برای فیدر هاو"""
        #self.f_header_frame = tk.Frame(self, bg=cream,height=25)
        #self.f_header_frame.grid(row=3, column=0, rowspan=1, columnspan=1, sticky='SNEW')
        #self.f_header_frame.grid_propagate(False)
        #font_label = ('Literal', 10, 'bold')
        #entry_bg = 'white'


        labels = [
            ("No", 0, 0,3),
            ("F.TYPE", 0, 1,14),
            ("POWER(KW)", 0, 2,10),
            ("PHASE", 0, 3,8),
            ("P.FACTOR", 0, 4,10),
            ("LENGTH(m)", 0, 5,10),
            ("CURRENT(A)", 0, 6,10),
            ("CABLE.SIZE", 0, 7,10),
            ("PH.NAME", 0, 8,10),
            ("DELTA V", 0, 9,10),
            ("BREAKER", 0, 10,10),
            ("BMetal(A)", 0, 11,10),
            ("SETTING(A)", 0, 12,10),
            ("CONTACTOR(A)", 0, 13,12), 
            ("FEEDER NAME", 0, 14,12),
        ]
        for text, r, c, w in labels:


            lbl = tk.Label(self.panel_top5, text=text,width=w, bg=cream, justify='center')
            lbl.grid(row=r, column=c,padx=2,pady=3,rowspan=1,columnspan=1,sticky='NW')
        


    def update_outputs(self):
        """به‌روزرسانی ویجت‌های خروجی."""
        self.total_power = self.panel_calculator.total_power  # دریافت مقدار از PanelCalculator
        self.max_current = self.panel_calculator.max_current  # دریافت مقدار از PanelCalculator
        self.demand_power = self.panel_calculator.demand_power  # دریافت مقدار از PanelCalculator
        self.demand_current = self.panel_calculator.demand_current  # دریافت مقدار از PanelCalculator
        self.main_cable_size = self.panel_calculator.main_cable_size  # دریافت مقدار از PanelCalculator
        self.main_breaker = self.panel_calculator.main_breaker  # دریافت مقدار از PanelCalculator
        self.upstream_panel_breaker = self.panel_calculator.upstream_panel_breaker  # دریافت مقدار از PanelCalculator
        self.p_delta_v = self.panel_calculator.p_delta_v
        self.panel_pf_var = self.panel_calculator.panel_pf_var  # دریافت مقدار از PanelCalculator
        self.main_cable_len = self.panel_calculator.main_cable_len
        self.panel_phase_name = self.panel_calculator.panel_phase_name  # دریافت مقدار از PanelCalculator

        self.out_total_power.config(text=f"{self.total_power:.2f}",background=green_fosfori)
        self.out_panel_current.config(text=f"{self.max_current:.1f}",background=green_fosfori)
        self.out_total_d_power.config(text=f"{self.demand_power:.2f}",background=green_fosfori)
        self.out_panel_d_current.config(text=f"{self.demand_current:.1f}",background=green_fosfori)
        self.out_panel_cable.config(text=f"{self.main_cable_size}",background=green_fosfori)
        self.out_panel_breaker.config(text=f"{self.main_breaker}",background=green_fosfori)
        self.out_panel_upstream_cb.config(text=f"{self.upstream_panel_breaker}",background=green_fosfori)
        self.out_panel_delta_v.config(text=f"{self.p_delta_v:.2f}",background=green_fosfori)
        self.out_total_PF.config(text=f"{self.panel_pf_var:.2f}",background=green_fosfori)
        self.panel_phase_name = self.panel_phase_input.get()
        self.panel_cable_len = int(self.main_cable_len_input.get())
        
    
    def get_tab_name_from_notebook(self):
        return self.master.tab(self, option="text")

        
    def reset_feeders(self):
        """ریست کردن تمام فیدرها"""
        for widget in self.panel_top5.winfo_children():
            widget.destroy()
        self.feeders.clear()
    
    def del_row(self):
        """حذف آخرین فیدر"""
        if self.feeders:
            last_feeder = self.feeders.pop()
            for widget in last_feeder.widgets:
                widget.destroy()
                

    
    def add_feeder(self):
        """اضافه کردن یک فیدر جدید."""
        row_idx = len(self.feeders)
        feeder = Feeder(self.panel_top5, row_idx, self.panel_calculator)
        self.feeders.append(feeder)
        self.panel_calculator.feeders = self.feeders  # بروزرسانی لیست فیدر در PanelCalculator
    
    def get_all_feeders_data(self):
        return [feeder.get_data() for feeder in self.feeders]

    def reset_feeders(self):
        """ریست کردن تمام فیدرها."""
        for feeder in self.feeders:
            feeder.destroy()
        self.feeders.clear()
        
    def calculate(self):
        self.panel_calculator.feeders = self.feeders
        self.panel_calculator.calculate()

            
    def load_data(self, panel_info, feeders_list):
        """بارگذاری داده‌ها در تب جدید"""

        # ست کردن اطلاعات ورودی
        self.panel_phase_input.set(panel_info["Phase"])
        self.panel_d_f_input.delete(0, tk.END)
        self.panel_d_f_input.insert(0, panel_info["Demand Factor"])

        self.upstream_panel_input.delete(0, tk.END)
        self.upstream_panel_input.insert(0, panel_info.get("Upstream Panel", ""))

        self.main_cable_len_input.delete(0, tk.END)
        self.main_cable_len_input.insert(0, panel_info["Cable Length(m)"])

        self.temp_input.set(panel_info["Temperature(c)"])
        self.installation_input.set(panel_info["Installation Type"])
        self.insulation_input.set(panel_info["Insulation"])
        self.max_vdrop_input.delete(0, tk.END)
        self.max_vdrop_input.insert(0, panel_info["Max Voltage Drop(%)"])
        # پاک کردن فیدرهای قبلی (اگر چیزی بود)
        self.reset_feeders()
        print("Feeders List Columns:", feeders_list.columns)
        # اضافه کردن فیدرهای جدید
        for idx, feeder in feeders_list.iterrows():
            self.add_feeder()
            feeder_obj = self.feeders[-1]
            try:
                feeder_obj.f_type_var.set(feeder["type"])
                feeder_obj.f_power_var.set(feeder["power"])
                feeder_obj.f_phase_var.set(feeder["phase"])
                feeder_obj.f_pf_var.set(feeder["pf"])
                feeder_obj.f_cable_len_var.set(feeder["cable_len"])
            except KeyError as e:
                messagebox.showerror("Error", f"Missing column in feeders_list: {e}")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while loading feeder data: {e}")


class ProjectInfoTab(ttk.Frame):
    def __init__(self, parent, project_name, window):
        super().__init__(parent)
        self.project_name = project_name.upper()  # Corrected to call upper() without parameters
        self.window = window
        self.panels_info = {}  # key: panel_name, value: panel_data
        self.panel_tabs = []  # لیست تمام PanelTab های ساخته شده 
        self.project_info = {}

        self.create_widgets()
        
    def add_row_to_active_tab(self):
        current_tab = self.notebook.nametowidget(self.notebook.select())
        if hasattr(current_tab, 'add_feeder'):
            current_tab.add_feeder()
    # [تمام متدهای مربوط به رابط کاربری ProjectInfoTab اینجا قرار میگیرد]
    # (create_widgets, add_panel, delete_active_tab, export_all_tabs_to_excel و...)
    def create_widgets(self):
        self.panel_win = tk.Toplevel(self.window)
        self.panel_win.title(f"{self.project_name} Project")
        self.panel_win.geometry("1300x850+100+50")
        self.panel_win.state("zoomed")
        # ایجاد فریم اصلی درون notebook
        self.header=tk.Frame(self.panel_win,bg=blue_dark,height=130,width=self.panel_win.winfo_screenwidth())
        self.header.grid(row=0,column=0,rowspan=1,columnspan=1,sticky='snew')
        
        self.under_header=tk.Frame(self.panel_win,bg=green2,height=2,width=self.panel_win.winfo_screenwidth())
        self.under_header.grid(row=1,column=0,rowspan=1,columnspan=1,sticky='snew')
        
        # ایجاد فریم برای نوت‌بوک

        self.notebook = ttk.Notebook(self.panel_win,height=620,width=self.panel_win.winfo_screenwidth())# ایجاد نوت‌بوک برای تب‌ها
        self.notebook.grid(row=3,column=0,rowspan=1,columnspan=1,sticky='snew') 
                 
        self.tab_control = ttk.Notebook(self.panel_win,height=620,width=self.panel_win.winfo_screenwidth())
        self.panel_list=[] # لیست برای ذخیره نام تب‌ها
        self.tab_data = {}  # برای ذخیره اطلاعات هر تب

        # ایجاد تب اطلاعات پروژه
        ptab = ttk.Frame(self.notebook,width=self.panel_win.winfo_screenwidth(),height=600)
        self.notebook.add(ptab, text='Project Info(' + self.project_name + ')')
        self.notebook.select(ptab)  # انتخاب تب اطلاعات پروژه به عنوان تب فعال
        # محتوای تب پروژه
        project_top0=tk.Frame(ptab,bg=blue_dark,border=2,height=100,width=1500) # پنجره panel_top0
        project_top0.grid_propagate(False)
        project_top0.grid(row=0,column=0,rowspan=1,columnspan=1,sticky='snew')
        
        project_top1=tk.Frame(ptab,bg=blue_dark,border=2,height=550,width=1500)           # پنجره panel_top1
        project_top1.grid_propagate(False)
        project_top1.grid(row=1,column=0,rowspan=1,columnspan=1,sticky='snew')
        
        # فیلدهای اطلاعات پروژه
        fields = [
            ("Project Name:", self.project_name, 24, False),
            ("Project Address:", "", 48, True),
            ("Client Name:", "", 24, True),
            ("Client Address:", "", 24, True),
            ("Designer Name:", "", 24, True),
            ("Designer Contact:", "", 24, True),
        ]

        lbl_project_prop = ttk.Label(project_top0, text="Project Information:", font=('Helvetica', 16))
        lbl_project_prop.grid(row=0, column=0, columnspan=2, pady=10)

        self.project_fields = {}  # دیکشنری برای ذخیره ویجت‌ها

        for idx, (label_text, default_value, entry_width, editable) in enumerate(fields, start=1):
            lbl = ttk.Label(project_top1, text=label_text, width=18, anchor='w')
            lbl.grid(row=idx, column=0, padx=5, pady=5, sticky='w')

            if editable:
                ent = ttk.Entry(project_top1, width=entry_width)
                ent.insert(0, default_value)
            else:
                ent = ttk.Label(project_top1, text=default_value, width=entry_width)

            ent.grid(row=idx, column=1, padx=5, pady=5, sticky='ew')

            # ذخیره ویجت برای دسترسی بعدی
            self.project_fields[label_text.strip(':')] = ent
            
        a = ("Helvetica", "10", "bold")
        c = ("Literal", "12", "bold")
        self.company_lbl=tk.Label(self.header,text="گروه نرم افزاری مانی نیروی البرز",justify='left',
                                  font=('Vazir','16','bold'),foreground="white",background=blue_dark)
        self.company_lbl.grid(row=0,column=0,padx=5,pady=5,sticky='snew')
        
        self.control_panel = tk.Frame(self.header, bg='white', width=self.panel_win.winfo_screenwidth(), height=50)
        self.control_panel.grid(row=1, column=0, padx=15, pady=5, sticky='snew')

        # دکمه‌های کنترل تب‌ها
        
        self_add_button = tk.Button(self.control_panel, text="Add Panel", width=20, font=a,
                                    background=blue_light, command=self.add_panel)
        self_add_button.grid(row=0,column=0,padx=5,pady=5,sticky='snew')
        
        
        self.excell_button = tk.Button(self.control_panel, text="Export To Excel", width=20, font=a,
                                       background=green4, foreground=white, command=ReportGenerator.export_all_tabs_to_excel)
        self.excell_button.grid(row=0, column=2, padx=5, pady=5, sticky='snew')

        btn_rename_tab = tk.Button(self.control_panel, text="Rename Panel", width=20, font=a,
                                   bg='#a3b18a', command=self.rename_panel_tab)
        btn_rename_tab.grid(row=0, column=3, padx=5)

        btn_delete_tab = tk.Button(self.control_panel, text="Delete Specific Panel", width=20, font=a,
                                   bg='#e63946', command=self.delete_specific_tab)
        btn_delete_tab.grid(row=0, column=4, padx=5)

        btn_export_tab = tk.Button(self.control_panel, text="Export Single Panel", width=20, font=a,
                                   bg='#4dff4d', command=ReportGenerator.export_single_panel)
        btn_export_tab.grid(row=0, column=5, padx=5)
        
        btn_duplicate_tab = tk.Button(self.control_panel, text="Duplicate Panel", width=20, font=a,
                                      bg='#6a994e', command=self.duplicate_panel_tab)
        btn_duplicate_tab.grid(row=0, column=6, padx=5)

    def rename_panel_tab(self):
        """باز کردن دیالوگ انتخاب تب برای تغییر نام"""
        def rename_action(selected_tab):
            new_name = simpledialog.askstring("Rename Panel", "Enter the new name:")
            new_name=new_name.upper()
            existing_names = [panel.panel_name for panel in self.panel_tabs]
            if new_name and self.validate_panel_name(new_name, existing_names):
                self.notebook.tab(selected_tab, text=new_name)
                
        self.select_panel_tab(rename_action)

    def delete_specific_tab(self):
        """باز کردن دیالوگ انتخاب تب برای حذف"""
        def delete_action(selected_tab):
            confirm = messagebox.askyesno("Confirm", "Are you sure you want to delete this panel?")
            if confirm:
                self.notebook.forget(selected_tab)
                self.panel_tabs.remove(selected_tab)
        self.select_panel_tab(delete_action)
        
    def duplicate_panel_tab(self):
        """دیالوگ انتخاب تب برای کپی کردن یک پنل"""
        def duplicate_action(selected_tab):
            try:
                # دریافت اطلاعات از تب انتخاب شده
                panel_data, feeders_data = selected_tab.get_output()

                # گرفتن اسم تب اصلی
                old_name = self.notebook.tab(selected_tab, "text")
                new_name = simpledialog.askstring("Duplicate Panel", "Enter name for duplicated panel:", initialvalue=f"{old_name}Copy")
                existing_names = [panel.panel_name for panel in self.panel_tabs]
                if not new_name or not self.validate_panel_name(new_name, existing_names):
                    return

                # ساختن تب جدید
                new_tab = PanelTab(self.notebook, new_name, project_info=self)
                self.notebook.add(new_tab, text=new_name)
                self.notebook.select(new_tab)
                self.panel_tabs.append(new_tab)

                # بارگذاری داده‌ها داخل تب جدید
                new_tab.load_data(panel_data.iloc[0], feeders_data)

                messagebox.showinfo("Success", "Panel duplicated successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to duplicate panel: {e}")

        self.select_panel_tab(duplicate_action)

    def add_panel(self):
        #for idx, panel_tab in enumerate(self.panel_tabs):
            #panel_tab.calculate() 
        panel_name = simpledialog.askstring("New Panel", "Enter panel name:")
        # فرض کنیم existing_names لیستی از نام تمام تب‌های فعلیه
        existing_names = [panel.panel_name for panel in self.panel_tabs]
        if panel_name and self.validate_panel_name(panel_name, existing_names):
            new_tab = PanelTab(self.notebook, panel_name, project_info=self)
            # اضافه کردن تب جدید به نوت‌بوک
            self.notebook.add(new_tab, text=panel_name)
            self.notebook.select(new_tab)
            self.panel_sheet = new_tab  # به‌روزرسانی مرجع
            
            self.panel_name = panel_name
            
            self.p_demand_factor = tk.DoubleVar()
            self.p_cable_len = tk.IntVar()
            self.p_temp = tk.IntVar()
            self.instalation_var = tk.StringVar()
            self.insulation_var = tk.StringVar()
            self.p_max_volage_drop = tk.DoubleVar()
            
            # اضافه کردن به لیست پنل‌ها
            self.panel_tabs.append(new_tab)
            
            # بررسی: آیا این تابلو پایین‌دستی کسی هست؟
            #upstream_name = new_tab.upstream_panel_input.get().strip()
            #for tab in self.panel_tabs:
            #    if tab.panel_name == upstream_name and tab != new_tab:
            #        # فیدر مربوطه در تابلوی بالادستی ایجاد شود
            #        if hasattr(new_tab, "panel_data") and new_tab.panel_data:
            #            self.create_feeder_from_downstream(tab, new_tab)
#
            ## بررسی: آیا کسی پایین‌دستی این تابلو هست؟
            #for tab in self.panel_tabs:
            #    if tab.upstream_panel_input.get().strip() == panel_name and tab != new_tab:
            #        if hasattr(tab, "panel_data") and tab.panel_data:
            #            self.create_feeder_from_downstream(new_tab, tab)

        
        global breakers,feeder_types
        breakers=(6,10,16,20,25,32,40,50,63,80,100,125,160,200,250,320,400,630,800,1000,1250,1600,2000)
        feeder_types= ("Lighting","Socket","Equipment","Motor(1P-DOL)","Motor(3P-DOL)","Motor(3P-YD)","Panel")
        self.feeders_data = []
        self.panel_data=[]
        self.total_power_var= 0
        self.current_list = [0, 0, 0]
        self.demand_current=0
        self.feeders=[]
        self.rows2=[]
        self.rows3=[]
        self.current_outs=[]
        self.cable_out=[]
        self.phase_n_out=[]
        self.delta_v_out=[]
        self.breaker_out=[]
       
    def update_project_list_display(self):

        self.project_list_display.delete(0, tk.END)
        for project in project_list:
            self.project_list_display.insert(tk.END, project)


        self.project_list_display = tk.Listbox(self.new_window, width=50, height=10)
        self.project_list_display.grid(row=2, column=0, padx=10, pady=10, sticky='nsew')
        
    def select_panel_tab(self, action_callback):
        """باز کردن دیالوگ انتخاب یک تب برای انجام عملیات خاص"""
        if not self.panel_tabs:
            messagebox.showerror("Error", "No panel tabs available.")
            return

        selection_win = tk.Toplevel(self.window)
        selection_win.title("Select a Panel")
        selection_win.geometry("300x250")
        selection_win.grab_set()

        lbl = tk.Label(selection_win, text="Select a Panel Tab:", font=("Helvetica", 12))
        lbl.pack(pady=10)

        listbox = tk.Listbox(selection_win, height=10)
        for idx, panel_tab in enumerate(self.panel_tabs):
            tab_text = self.notebook.tab(panel_tab, option="text")
            listbox.insert(tk.END, f"{idx+1}: {tab_text}")
        listbox.pack(pady=10, fill='both', expand=True)

        def on_select():
            selected_index = listbox.curselection()
            if selected_index:
                index = selected_index[0]
                selected_tab = self.panel_tabs[index]
                action_callback(selected_tab)
                selection_win.destroy()
            else:
                messagebox.showwarning("Warning", "Please select a panel.")

        btn_ok = tk.Button(selection_win, text="OK", command=on_select)
        btn_ok.pack(pady=5)


    def validate_panel_name(self,name, existing_names):
        """
        بررسی معتبر بودن نام پنل (تابلو).

        پارامترها:
            name (str): نام واردشده توسط کاربر.
            existing_names (set or list): مجموعه‌ای از نام‌های موجود برای بررسی تکراری بودن.

        خروجی:
            True اگر نام معتبر بود، False در غیر اینصورت.
        """
        name = name.strip()

        # فقط حروف و اعداد انگلیسی
        if not re.fullmatch(r'[A-Za-z0-9]+', name):
            messagebox.showerror("Invalid Name", "Only English letters and digits are allowed.\nNo spaces or special characters.")
            return False

        # بررسی تکراری بودن
        if name in existing_names:
            messagebox.showerror("Duplicate Name", f"The panel name '{name}' already exists.")
            return False

        return True


class PanelProject:
    def __init__(self,parent, root):
        self.parent = parent  # ذخیره مرجع CreateProject
        self.window=root

        project_name = simpledialog.askstring("New Tab", "Enter the name for the Project")
        if not project_name:
            messagebox.showerror("Error", "Please Insert a Name for Project")
            return
        self.project_info_tab = ProjectInfoTab(self.window, project_name, self.window)


    def on_tab_change(self, event):
        selected_tab = self.notebook.tab(self.notebook.select(), "text")
        print(f"Selected tab: {selected_tab}")
        
class CreateProject:
    def __init__(self,root):
        self.root = root
        self.language = "En.Language"
        self.project_name = simpledialog.askstring("New Tab", "Enter the name for the Project")
        if not self.project_name:
            messagebox.showerror("Error", "Please Insert a Name for Project")
            return
        self.setup_ui2()
      
    def setup_ui2(self):
        """Initializes the main user interface."""
        self.root = tk.Toplevel(self.root)
        self.root.title("Panel Project")
        self.root.geometry("1300x850+100+50")
        self.root.state("zoomed")
        self.root.configure(bg=COLORS["blue_dark"])
        self.create_label_entry()
        self.create_main_buttons2()

        
     
    def create_label_entry(self):
        lbl = tk.Label(self.root, text=self.project_name, font=('Corbel', '16', 'bold'))
        lbl.grid(row=0, column=0, padx=10, pady=10)


        

    def create_main_buttons2(self):
        """Creates the main navigation buttons."""
        button_config = {
            "width": 30,
            "height": 2,
            "font": ("Corbel", "12", "bold"),
            "background": COLORS["green4"],
            "foreground": COLORS["white"],
            "activebackground": COLORS["green2"],
            "activeforeground": COLORS["blue_dark"],
            "highlightthickness": 2,
            "borderwidth": 4,
            "relief": tk.RAISED,
            "highlightbackground": COLORS["blue_dark"],
            "highlightcolor": COLORS["blue_light"],
        }

        buttons = [
            ("PANEL PROJECT", self.open_panel_project),
            ("UPS BATTERY AMPERE CALCULATION", self.open_ups_battery_ampere),
            ("UPS SUPPORT TIME CALCULATION",self.open_ups_support_time),
            ("PANEL PROJECT", self.open_panel_project),
            ("UPS BATTERY AMPERE CALCULATION", self.open_ups_battery_ampere),
            ("UPS SUPPORT TIME CALCULATION",self.open_ups_support_time),
        ]

        for i, (text, command) in enumerate(buttons):
            button = tk.Button(self.root, text=text, command=command, **button_config)
            button.grid(row=i+7, column=0, pady=5, padx=5)
        
            
    def open_panel_project(self):
        """Opens the Panel Project window."""
        self.window = self.root
        self.panel_project = PanelProject(self, self.root)  # ارسال project_info به PanelProject

    def open_ups_battery_ampere(self):
        """Opens the UPS Battery Ampere Calculation window."""
        UPSBatteryAmpere(self.root)

    def open_ups_support_time(self):
        """Opens the UPS Support Time Calculation window."""
        UPSBatterySupport(self.root)
        
