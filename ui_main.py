import tkinter as tk
from tkinter import ttk, Menu

from utilities import COLORS
from ups_calculations import UPSBatteryAmpere, UPSBatterySupport
from panel_project import PanelProject

class MainUI:
    """Main application class."""
    
    def __init__(self, root):
        self.root = root
        self.language = "En.Language"
        self.setup_ui()

    def setup_ui(self):
        """Initializes the main user interface."""
        self.root.title("VEBER Electrical Calculation")
        self.root.geometry("1050x450+100+50")
        self.root.configure(bg=COLORS["blue_dark"])

        self.create_menu_bar()

        self.create_main_buttons()

        company_lbl = tk.Label(
            self.root,
            text="Mani Niroo Company",
            font=('Corbel', '16', 'bold'),
            foreground=COLORS["white"],
            background=COLORS["blue_dark"]
        )
        company_lbl.grid(row=2, column=4, pady=5, padx=5)

    def create_menu_bar(self):
        """Creates the main menu bar."""
        menubar = Menu(self.root)

        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="New", command=lambda: None)
        filemenu.add_command(label="Open", command=lambda: None)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=filemenu)

        editmenu = Menu(menubar, tearoff=0)
        editmenu.add_command(label="Change Language", command=self.change_language)
        menubar.add_cascade(label="Edit", menu=editmenu)

        self.root.config(menu=menubar)

    def create_main_buttons(self):
        """Creates the main navigation buttons."""
        button_config = {
            "width": 38,
            "height": 3,
            "font": ("Corbel", "12", "bold"),
            "background": COLORS["green4"],
            "foreground": COLORS["white"]
        }

        buttons = [
            ("PANEL CALCULATION Software", self.open_panel_project),
            ("UPS SUPPORT TIME CALCULATION", self.open_ups_support_time),
            ("UPS BATTERY AMPERE CALCULATION", self.open_ups_battery_ampere)
        ]

        for i, (text, command) in enumerate(buttons):
            button = tk.Button(self.root, text=text, command=command, **button_config)
            button.grid(row=i, column=0, pady=5, padx=5)

    def change_language(self):
        """Changes the language of the UI."""
        if self.language == "En.Language":
            self.language = "Per.Language"
            print("Language switched to Persian.")
        else:
            self.language = "En.Language"
            print("Language switched to English.")

    def open_panel_project(self):
        """Opens the Panel Project Management window."""
        PanelProject(self.root)

    def open_ups_battery_ampere(self):
        """Opens the UPS Battery Ampere Calculation window."""
        UPSBatteryAmpere(self.root)

    def open_ups_support_time(self):
        """Opens the UPS Support Time Calculation window."""
        UPSBatterySupport(self.root)