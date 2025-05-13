import pandas as pd
import openpyxl
from xlsxwriter import Workbook
from tkinter import ttk, simpledialog, messagebox, filedialog#             "borderwidth": 0,

class PanelCalculator:
    def __init__(self, panel_phase, demand_factor, main_cable_len,
                 temp, installation, insulation, max_vdrop,project_info):
        self.panel_phase_name = panel_phase
        self.demand_factor = demand_factor
        self.main_cable_len = main_cable_len
        self.temp = temp
        self.installation = installation  # مقداردهی installation
        self.insulation = insulation
        self.max_vdrop = max_vdrop
        self.project_info = project_info  # مقداردهی project_info
        self.sheets = self.load_excel_data()
        self.feeders = []
        self.max_current = 0  # مقداردهی اولیه
        print(self.panel_phase_name)

    def load_excel_data(self):
        """بارگذاری داده‌ها از فایل اکسل."""
        try:
            file_path = "data.xlsx"
            self.sheets = {
                "In Air": pd.read_excel(file_path, sheet_name="In Air"),
                "In Ground": pd.read_excel(file_path, sheet_name="In Ground"),
                "m3_yd": pd.read_excel(file_path, sheet_name="m3_yd"),
                "m3_dol": pd.read_excel(file_path, sheet_name="m3_dol"),
                "m1_dol": pd.read_excel(file_path, sheet_name="m1_dol"),
                "current_table": pd.read_excel(file_path, sheet_name="current_table"),
                "lighting_table": pd.read_excel(file_path, sheet_name="lighting_table"),
                "socket_table": pd.read_excel(file_path, sheet_name="socket_table"),
                "breaker": pd.read_excel(file_path, sheet_name="breaker"),
                "panel_table": pd.read_excel(file_path, sheet_name="panel_table")
            }
            return self.sheets
        except FileNotFoundError:
            messagebox.showerror("Error", "The file 'data.xlsx' was not found.")
            raise
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while loading data: {e}")
            raise
        
    def validate_inputs(self):
        try:
            #errors = []
#
            #if not self.panel_phase_input.get():
            #    errors.append("Panel Phase")
            #if not self.panel_d_f_input.get().strip():
            #    errors.append("Demand Factor")
            #if not self.main_cable_len_input.get().strip():
            #    errors.append("Main Cable Length(m)")
            #if not self.temp_input.get():
            #    errors.append("Ambient Temperature(c)")
            #if not self.installation_input.get():
            #    errors.append("Installation Type")
            #if not self.insulation_input.get():
            #    errors.append("Insulation Type")
            #if not self.max_vdrop_input.get().strip():
            #    errors.append("Max Voltage Drop(%)")
#
            #if errors:
            #    raise ValueError("لطفاً فیلدهای زیر را تکمیل کنید:\n" + "\n".join(errors))

            #self.panel_phase_name = self.panel_phase_input.get()
            self.panel_cable_len = int(self.main_cable_len_input.get())
            self.panel_d_f = float(self.panel_d_f_input.get())
            self.temp = int(self.temp_input.get())
            self.installation_type = self.installation_input.get()
            self.insulation = self.insulation_input.get()
            self.max_voltage_drop = float(self.max_vdrop_input.get())
            

            if not (0 < self.panel_d_f <= 1):
                raise ValueError("مقدار Demand Factor باید بین 0 و 1 باشد.")

        except ValueError as e:
            messagebox.showerror("خطای ورودی", str(e))
            raise
        
    def calculate(self):      
        try:

            print("2")
            self.load_excel_data()
            print("3")
            k = self.calculate_k_factor()# محاسبه مقادیر اولیه
            print(f"K Factor: {k}")
            self.calculate_feeders()  # محاسبه فیدرها
            print("5")
            self.calculate_panel(k)# محاسبه مقادیر پنل
            print("6")
            self.update_outputs()  # به‌روزرسانی خروجی‌ها
        except FileNotFoundError:
            messagebox.showerror("Error", "The file 'data.xlsx' was not found.")
        
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        # اگر parent یک ProjectInfoTab است، بهش بگو برو فیدرهای بالادستی منو آپدیت کن
        if hasattr(self.project_info, "update_feeders_connected_to"):
            self.project_info.update_feeders_connected_to(self.panel_name)
            self.feeders = self.project_info.feeders if hasattr(self.project_info, 'feeders') else []
        if hasattr(self.project_info, "update_outputs"):
            self.project_info.update_outputs()
            
        try:
            self.total_power = sum(feeder["power"] for feeder in self.feeders_data) 
            print(f"Total Power (KW): {self.total_power}")
            self.project_info.total_power = self.total_power  # انتقال مقدار به PanelTab
            print("انتقال total_power اطلاعات پروژه")
            self.max_current = max(feeder["current"] for feeder in self.feeders_data)  # محاسبه بیشترین جریان
            print(f" max_current  {self.max_current}")
            self.project_info.max_current = self.max_current
            print("انتقال max_current اطلاعات پروژه")
            self.project_info.update_outputs()
            print("انتفال خروجی ها به پنل تب")
        except Exception as e:
            raise ValueError(f"Error in calculation: {e}")
        
        """محاسبه مقادیر پنل و به‌روزرسانی ویجت‌ها."""
    def update_outputs(self):
        """به‌روزرسانی خروجی‌های رابط کاربری."""
        try:
            # به‌روزرسانی مقادیر پنل
            if hasattr(self, "panel_data"):
                print("Updating panel outputs...")
                print(f"Panel Name: {self.panel_data.get('Panel Name', '')}")
                print(f"Total Power (KW): {self.panel_data.get('Total Power(KW)', '')}")
                print(f"Demand Current (A): {self.panel_data.get('Demand Current(A)', '')}")
                print(f"Main Cable Size (mm2): {self.panel_data.get('Main Cable Size(mm2)', '')}")
                print(f"Panel Breaker (A): {self.panel_data.get('Panel Breaker(A)', '')}")
            
            # به‌روزرسانی مقادیر فیدرها
            if hasattr(self, "feeders_data"):
                print("Updating feeder outputs...")
                for feeder in self.feeders_data:
                    print(f"Feeder Type: {feeder['type']}, Power: {feeder['power']} KW, Current: {feeder['current']} A")
        except Exception as e:
            print(f"Error updating outputs: {e}")
    #def update_feeder_data(self):
    def calculate_k_factor(self):
        #محاسبه فاکتور K بر اساس دما و نوع عایق.
        print(f"Calculating K factor for installation type: {self.installation}")
        try:
            # بررسی اینکه installation مقداردهی شده است
            if not self.installation:
                raise ValueError("Installation type is not set.")

            # دریافت شیت مربوط به installation
            kf = self.sheets.get(self.installation)
            if kf is None:
                raise ValueError(f"Sheet for installation type '{self.installation}' not found.")

            # بررسی وجود ستون‌های موردنیاز
            if "TEMP" not in kf.columns or self.insulation not in kf.columns:
                raise ValueError(f"Required columns ('TEMP' or '{self.insulation}') not found in sheet '{self.installation}'.")

            # جستجوی مقدار K بر اساس دما و نوع عایق
            temp_values = kf[kf["TEMP"] == int(self.temp)]
            if temp_values.empty:
                raise ValueError(f"No matching temperature '{self.temp}' found in sheet '{self.installation}'.")

            k_value = float(temp_values[self.insulation].values[0])
            return k_value

        except ValueError as e:
            messagebox.showerror("Error", str(e))
            raise
        except Exception as e:
            raise ValueError(f"An unexpected error occurred while calculating K factor: {e}")

        
        
        #self.sheets = self.load_excel_data()
        #print("sheets ok")
        #try:
        #    kf = self.sheets.get(self.installation)
        #    print(kf)
        #    if kf is None:
        #        raise ValueError(f"Sheet for installation type '{self.installation}' not found.")
#
        #    if "TEMP" not in kf.columns or self.insulation not in kf.columns:
        #        raise ValueError(f"Required columns ('TEMP' or '{self.insulation}') not found in sheet '{self.installation}'.")
#
        #    # جستجوی مقدار K
        #    k_value = float(kf[kf["TEMP"] == int(self.temp)][self.insulation].values[0])
        #    return k_value
        #except (IndexError, KeyError):
        #    raise ValueError("Invalid temperature or insulation type.")
        #except Exception as e:
        #    raise ValueError(f"An error occurred while calculating K factor: {e}")
    
    

    def calculate_feeders(self):
        
        if not hasattr(self, 'sheets') or not self.sheets:
            self.load_excel_data()
            print("excel data loaded successfully.")
        if not self.feeders:
            print("No feeders available.")
            return

        try:
            self.breaker = self.sheets["breaker"]
            self.lighting_table = self.sheets["lighting_table"]
            self.socket_table = self.sheets["socket_table"]
            self.m1_dol = self.sheets["m1_dol"]
            self.m3_dol = self.sheets["m3_dol"]
            self.m3_yd = self.sheets["m3_yd"]
            self.panel_table = self.sheets["panel_table"]
        except KeyError as e:
            messagebox.showerror("Error", f"Missing sheet: {e}")
            return
        print("excel data loaded successfully2.")
        self.feeders_data = []
        self.current_list = [0, 0, 0]
        self.total_power = 0

        for i, feeder in enumerate(self.feeders):
            data = feeder.get_data()
            if not data["power"] or not data["phase"] or not data["pf"] or not data["cable_len"]:
                continue
            
            f_type = data["type"]
            f_power = float(data["power"])
            f_phase = int(data["phase"])
            f_pf = float(data["pf"])
            f_len = float(data["cable_len"])
            print("data loaded")
            current = (f_power * 1000) / (f_phase * f_pf * 230)
            
            print(f"Current: {current}")
            
            self.total_power += f_power
            print("total power in feeder calculated")
            f_number = f"F{i+1}"

            if self.panel_phase_name == "RST" and f_phase == 1:
                idx = self.current_list.index(min(self.current_list))
                self.current_list[idx] += current
                f_phase_name = ["R", "S", "T"][idx]
            else:
                idx = {"R": 0, "S": 1, "T": 2}.get(self.panel_phase_name, 0)
                for j in range(f_phase):
                    self.current_list[(idx + j) % 3] += current
                    f_phase_name = self.panel_phase_name
            print("phase name in feeder calculated")
            cable, breaker, bmetal, setting, contactor = self.select_equipment(f_type, f_phase, f_power, current)
            print("equipment in feeder calculated")
            f_delta_v = self.calc_delta_v_feeder(f_power, f_len, cable, f_phase)
            print("delta v in feeder calculated")
            self.feeders_data.append({
                "type": f_type, "power": f_power, "phase": f_phase, "pf": f_pf, "cable_len": f_len,
                "current": current, "cable": cable, "phase_name": f_phase_name, "delta_v": f_delta_v,
                "breaker": breaker, "number": f_number,
                "bmetal": bmetal, "setting": setting, "contactor": contactor
            })
            
    def select_equipment(self, f_type, f_phase, f_power, current):
        """انتخاب کابل و کلید بر اساس نوع فیدر"""
        table = None
        bmetal = setting = contactor = "-"
        cb_table = self.breaker
        current_cb = cb_table[cb_table['c_breaker'] >= current * 1.25].iloc[0][1]
    
        if f_type == "Lighting":
            table = self.lighting_table
        elif f_type == "Socket" or f_type == "Equipment":
            table = self.socket_table
        elif f_type == "Motor(1P-DOL)":
            table = self.m1_dol
        elif f_type == "Motor(3P-DOL)":
            table = self.m3_dol
        elif f_type == "Motor(3P-YD)":
            table = self.m3_yd
        #elif f_type == "Panel":
            #table = self.panel_table   
    
        if table is not None:
            if "Motor" in f_type:
                row = table[table['POWER'] >= f_power].iloc[0]
                cable = row[5]
                bmetal, setting, contactor = row[3], row[2], row[4]
                current_cb = row[1]
            else:
                col = "1PHASE_A" if f_phase == 1 else "3PHASE_A"
                cable = table[table[col] >= current_cb].iloc[0,5]
        else:
            cable = "-"
    
        return cable, current_cb, bmetal, setting, contactor
    
    def calc_delta_v_feeder(self, power, length, cable, phase):
        """محاسبه افت ولتاژ یک فیدر"""
        try:
            cable = float(cable)
            if phase == 1:
                return round((power * 1000 * length * 2) / (56 * cable * 230), 2)
            else:
                return round((power * 1000 * length) / (56 * cable * 400), 2)
        except:
            return "-"
            
    def f_current_calculator(self, f_power, f_phase, f_pf, ):
        """محاسبه جریان فیدر بر اساس قدرت، فاز و ضریب قدرت."""
        current = (f_power * 1000) / (f_phase * f_pf * 230)
        return current 
    
    def select_f_equipment(self, f_type, f_power, current):
        """انتخاب کابل و کلید بر اساس نوع فیدر"""
        table = None
        bmetal = setting = contactor = "-"
        cb_table = self.breaker
        current_cb = cb_table[cb_table['c_breaker'] >= current * 1.25].iloc[0, 1]
    
        if f_type == "Lighting":
            table = self.lighting_table
        elif f_type == "Socket" or f_type == "Equipment":
            table = self.socket_table
        elif f_type == "Motor(1P-DOL)":
            table = self.m1_dol
        elif f_type == "Motor(3P-DOL)":
            table = self.m3_dol
        elif f_type == "Motor(3P-YD)":
            table = self.m3_yd
        #elif f_type == "Panel":
            #table = self.panel_table   
    
        if table is not None:
            if "Motor" in f_type:
                row = table[table['POWER'] >= f_power].iloc[0]
                bmetal, setting, contactor = row[3], row[2], row[4]
                current_cb = row[1]
            else:
                current_cb == "-",  bmetal=="-", setting=="-", contactor == "-"
        return  current_cb, bmetal, setting, contactor
            
    def f_cable_calculator(self, f_type, f_phase, f_power, f_current):
        """محاسبه سایز کابل بر اساس قدرت، فاز و ضریب قدرت."""
        table = None
        
        cb_table = self.breaker
        current_cb = cb_table[cb_table['c_breaker'] >= f_current * 1.25].iloc[0, 1]

        if f_type == "Lighting":
            table = self.lighting_table
        elif f_type == "Socket" or f_type == "Equipment":
            table = self.socket_table
        elif f_type == "Motor(1P-DOL)":
            table = self.m1_dol
        elif f_type == "Motor(3P-DOL)":
            table = self.m3_dol
        elif f_type == "Motor(3P-YD)":
            table = self.m3_yd
        #elif f_type == "Panel":
            #table = self.panel_table   
    
        if table is not None:
            if "Motor" in f_type:
                row = table[table['POWER'] >= f_power].iloc[0]
                cable = row[5]
            else:
                col = "1PHASE_A" if f_phase == 1 else "3PHASE_A"
                cable = table[table[col] >= current_cb].iloc[0, 5]
        else:
            cable = "-"
    
        return cable, current_cb
    
        
    def f_delta_v_calculator(self):
        """محاسبه افت ولتاژ یک فیدر"""
        self.cable = float(self.cable)
        try:
            def f_delta_v_calc():
                if self.f_phase == 1:
                    return round((self.f_power * 1000 * self.f_cable_len * 2) / (56 * self.cable * 230), 2)
                else:
                    return round((self.f_power * 1000 * self.f_cable_len) / (56 * self.cable * 400), 2)
            self.f_delta_v = f_delta_v_calc()
            i = 0
            while self.f_delta_v > 3:
                i += 1
                if i >= len(self.filtered_table):
                    raise ValueError("No suitable cable size found for the given voltage drop.")
                self.cable = self.filtered_table.iloc[i]["SIZE"]
                self.f_delta_v = f_delta_v_calc()

            return self.f_delta_v
        except KeyError as e:
            raise ValueError(f"Missing column in filtered_table: {e}")
        except IndexError:
            raise ValueError("Index out of range while accessing cable size.")
        except Exception as e:
            raise ValueError(f"An unexpected error occurred while calculating Delta V: {e}")

    
    def calculate_panel(self, k):
        """محاسبه مقادیر مربوط به پنل."""

        try:
            
            self.total_power = sum(f["power"] for f in self.feeders_data)
            self.max_current=float(max(self.current_list))
            self.demand_current = self.max_current * self.panel_d_f
            self.demand_power = self.total_power * self.panel_d_f
            # محاسبه کابل اصلی
            self.derated_current = self.demand_current / k
            self.main_cable_size = self.get_main_cable_size()
            # محاسبه قدرت ظاهری و ضریب قدرت
            if self.panel_phase_name == "RST":
                self.panel_pf_var = self.demand_power * 1000 / (230 * 3 * self.demand_current)
            else:
                self.panel_pf_var = self.demand_power * 1000 / (230 * self.demand_current)

            self.main_breaker = self.get_main_breaker()
            self.upstream_panel_breaker = self.get_upstream_breaker()
            self.panel_delta_v_var = self.calc_delta_v()
            # مقداردهی panel_data
            self.panel_data = {
                "Panel Name": self.panel_name,
                "Phase": self.panel_phase_name,
                "Cable Length(m)": self.panel_cable_len,
                "Demand Factor": self.panel_d_f,
                "Temperature(c)": self.temp,
                "Installation Type": self.installation_type,
                "Insulation": self.insulation,
                "Max Voltage Drop(%)": self.max_voltage_drop,
                "Total Power(KW)": self.total_power,
                "Max Current(A)": self.max_current,
                "Demand Power(KW)": self.demand_power,
                "Demand Current(A)": self.demand_current,
                "Main Cable Size(mm2)": self.main_cable_size,
                "Panel Breaker(A)": self.main_breaker,
                "Upstream Breaker(A)": self.upstream_panel_breaker,
                "Delta V%": self.panel_delta_v_var,
                "Power Factor": self.panel_pf_var,
            }


        except Exception as e:
            raise ValueError(f"An error occurred while calculating panel: {e}")
        
    def get_main_cable_size(self):
        """محاسبه سایز کابل اصلی."""
        try:
            if self.panel_phase_name == "RST":
                f_phase = "3PHASE_A" if self.installation_type == "In Air" else "3PHASE_G"
            else:
                f_phase = "1PHASE_A" if self.installation_type == "In Air" else "1PHASE_G"

            current_table = self.sheets["current_table"]
            self.filtered_table = current_table[current_table[f_phase] >= self.derated_current]
            if self.filtered_table.empty:
                raise ValueError("No suitable cable size found for the given current.")
            return self.filtered_table.iloc[0]["SIZE"]
        except KeyError as e:
            raise ValueError(f"Missing column in current_table: {e}")
        except Exception as e:
            raise ValueError(f"An error occurred while getting main cable size: {e}")
    
    def get_main_breaker(self):
        try:
            filtered_breaker = self.breaker[self.breaker['c_breaker'] >= self.demand_current * 1.25]
            if filtered_breaker.empty:
                raise ValueError("No suitable breaker found for the given current.")
            self.main_breaker = filtered_breaker.iloc[0][1]
            return self.main_breaker
        except KeyError as e:
            raise ValueError(f"Missing column in breaker table: {e}")
        except Exception as e:
            raise ValueError(f"An error occurred while getting main breaker: {e}")

    def get_upstream_breaker(self):
        filtered_upstream_cb = self.breaker[self.breaker['c_breaker'] == self.main_breaker]
        upstream_cb = filtered_upstream_cb.iloc[0][2]
        return upstream_cb

    def calc_delta_v(self):
        try:
            def calculate_delta_v():
                if self.panel_phase_name == "RST":
                    return (self.demand_power * 1000 * self.panel_cable_len) / (56 * self.main_cable_size * 400)
                else:
                    return (self.demand_power * 1000 * self.panel_cable_len * 2) / (56 * self.main_cable_size * 230)

            self.p_delta_v = calculate_delta_v()
            i = 0
            while self.p_delta_v > self.max_voltage_drop:
                i += 1
                if i >= len(self.filtered_table):
                    raise ValueError("No suitable cable size found for the given voltage drop.")
                self.main_cable_size = self.filtered_table.iloc[i]["SIZE"]
                self.p_delta_v = calculate_delta_v()

            return self.p_delta_v

        except KeyError as e:
            raise ValueError(f"Missing column in filtered_table: {e}")
        except IndexError:
            raise ValueError("Index out of range while accessing cable size.")
        except Exception as e:
            raise ValueError(f"An unexpected error occurred while calculating Delta V: {e}")

    
    def get_output(self):
        """بازگرداندن اطلاعات پنل و فیدرها به صورت DataFrame."""
        
        self.calculate()
        try:
            panel_df = pd.DataFrame([self.panel_data])  # تبدیل panel_data به DataFrame
            print("Feeders Data:", self.feeders_data)
            feeders_df = pd.DataFrame(self.feeders_data)  # تبدیل feeders_data به DataFrame
            print(panel_df,feeders_df)
            return panel_df, feeders_df
        
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while getting output: {e}")
            return None, None
        
    def save_panel(self):
        """ذخیره اطلاعات پنل در فایل اکسل."""
        try:
            file_path = "panel_data.xlsx"
            data = {
                "Phase": self.panel_phase_name,
                "Cable Length(m)": self.panel_cable_len,
                "Demand Factor": self.panel_d_f,
                "Temperature(c)": self.temp,
                "Installation Type": self.installation_type,
                "Insulation": self.insulation,
                "Max Voltage Drop(%)": self.max_voltage_drop,
                
                "Total Power(KW)": self.total_power,
                "Max Current(A)": self.max_current,
                "Demand Power(KW)": self.demand_power,
                "Demand Current(A)": self.demand_current,
                
                "Main Cable Size(mm2)": self.main_cable_size,
                "Panel Breaker(A)": self.main_breaker,
                "Upstream Breaker(A)": self.upstream_panel_breaker,
                "Delta V%": self.p_delta_v,
                "Power Factor": self.panel_pf_var,

            }
            df = pd.DataFrame([data])
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", "Panel data saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving data: {e}")

class ReportGenerator:
    def __init__(self):
        self.formats = self.define_excel_formats()
        

    #def define_excel_formats(self):
         # [کدهای فرمت‌دهی اکسل]
         
        
    def format_excel_sheet(self,writer, sheet_name, df, start_row=0):
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        # --- فرمت‌ها ---
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#B7DEE8',  # آبی روشن
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1
        })
        cell_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1
        })
        float_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1,
            'num_format': '0.00'  # فقط دو رقم اعشار
        })
        # --- فرمت‌دهی هدرها ---
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(start_row, col_num, value, header_format)
            # تعیین حداکثر عرض ستون‌ها
            max_len = max(df[value].astype(str).map(len).max(), len(value)) + 2
            column_width = min(max_len, 15)
            worksheet.set_column(col_num, col_num, column_width)
        # --- فرمت‌دهی سلول‌ها ---
        for row_idx, row in df.iterrows():
            for col_idx, val in enumerate(row):
                cell_val = val
                if isinstance(val, float):
                    worksheet.write(row_idx + start_row + 1, col_idx, cell_val, float_format)
                else:
                    worksheet.write(row_idx + start_row + 1, col_idx, str(cell_val), cell_format)
    def format_project_info_sheet(self,writer, sheet_name):
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        # فرمت عمومی سلول‌ها
        cell_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1
        })
        # تعیین عرض ستون‌ها
        worksheet.set_column(0, 0, 24, cell_format)   # ستون اول: کلیدها
        worksheet.set_column(1, 1, 100, cell_format)  # ستون دوم: مقادیر

    #def generate_report(self, data):
        # [کدهای تولید گزارش]
        
    def export_single_panel(self):
        """باز کردن دیالوگ انتخاب تب برای خروجی گرفتن اکسل"""
        def export_action(selected_tab):
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if save_path:
                try:
                    panel_df, feeders_df = selected_tab.get_output()
                    sheet_name = selected_tab.panel_name  # استخراج نام تب به‌صورت رشته
                    writer = pd.ExcelWriter(save_path, engine='xlsxwriter')
                    panel_df.to_excel(writer, sheet_name=sheet_name, startrow=0, index=False)
                    self.format_excel_sheet(writer, sheet_name, panel_df, start_row=0)
                    feeders_df.to_excel(writer, sheet_name=sheet_name, startrow=len(panel_df)+3, index=False)
                    self.format_excel_sheet(writer, sheet_name, feeders_df, start_row=len(panel_df)+3)
                    writer.close()
                    messagebox.showinfo("Success", "Panel exported successfully!")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to export panel: {e}")
        self.select_panel_tab(export_action)
        
    def export_all_tabs_to_excel(self):

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Excel File"
        )
        #if not file_path.endswith(".xlsx"):
        #    file_path += ".xlsx"
        if not save_path:
            return
    
        writer = pd.ExcelWriter(save_path, engine='xlsxwriter')
        

        
                # --- Sheet 1: Project Info ---
        used_names = set()
        project_info = {}
        for key, widget in self.project_fields.items():
            if isinstance(widget, ttk.Entry):
                value = widget.get()
            elif isinstance(widget, ttk.Label):
                value = widget.cget("text")
            else:
                value = ""
            project_info[key] = value
        
        project_df = pd.DataFrame.from_dict(project_info, orient='index', columns=["Value"])
        project_df.to_excel(writer, sheet_name="Project Info")
        self.format_project_info_sheet(writer, "Project Info")
        
        for idx, panel_tab in enumerate(self.panel_tabs):
            panel_tab.calculate() 
            if hasattr(panel_tab, "get_output"): 
                try:
                    panel_df, feeders_df = panel_tab.get_output()
                    base_name = panel_tab.get_tab_name_from_notebook()
                    sheet_name = base_name
                    counter = 1
                    while sheet_name in used_names:
                        sheet_name = f"{base_name}_{counter}"
                        counter += 1
                    used_names.add(sheet_name)

                    # Exporting the data to the respective sheet
                    panel_df.to_excel(writer, sheet_name=sheet_name, startrow=0, index=False)
                    self.format_excel_sheet(writer, sheet_name, panel_df, start_row=0)
                    feeders_df.to_excel(writer, sheet_name=sheet_name, startrow=len(panel_df)+3, index=False)
                    self.format_excel_sheet(writer, sheet_name, feeders_df, start_row=len(panel_df)+3)
                except Exception as e:
                    print(f"Error exporting panel {idx+1}: {e}")
        
        

                # --- Sheet 2: Panels Summary ---
        summary_data = []

        for tab in self.panel_tabs:
            
            if hasattr(tab, "panel_data") and tab.panel_data:
                summary_data.append({
                    "Panel Name": tab.panel_name,
                    "Upstream": tab.upstream_panel_input.get(),
                    "Demand Power (kW)": tab.panel_data.get("Demand Power(KW)", ""),
                    "Phase": tab.panel_data.get("Phase", ""),
                    "Demand Current (A)": tab.demand_current,
                    "Power Factor": tab.panel_data.get("Power Factor", ""),
                    "Main Cable(mm2)": tab.panel_data.get("Main Cable Size(mm2)", ""),
                    "Cable Length(m)": tab.panel_data.get("Cable Length(m)", ""),
                    "Main Breaker(A)": tab.panel_data.get("Panel Breaker(A)", "")
                })
    
        if summary_data:
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name="Panels Summary", index=False)
            self.format_excel_sheet(writer, "Panels Summary", summary_df, start_row=0)


        writer.close()
        messagebox.showinfo("Export Complete", "All panels exported successfully!")
    ##################################################################################  
   