
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from ups_battery_ampere import UPSBatteryAmpere
from ups_support_time import UPSBatterySupport
from gui import CreateProject
from utilities import COLORS

class Application:
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
            text="مانی نیروی البرز",
            font=('Corbel', '16', 'bold'),
            foreground=COLORS["white"],
            background=COLORS["blue_dark"]
        )
        company_lbl.grid(row=0, column=1, pady=5, padx=5)

    def create_menu_bar(self):
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

    def create_main_buttons(self):
        """Creates the main navigation buttons."""
        button_config = {
            "width": 30,
            "height": 3,
            "font": ("Corbel", "14", "bold"),
            "background": COLORS["green4"],
            "foreground": COLORS["white"],
            "activebackground": COLORS["green2"],
            "activeforeground": COLORS["blue_dark"],
            "highlightthickness": 2,
            "borderwidth": 6,
            "highlightcolor": COLORS["blue_light"],
            "relief": tk.RAISED,
            "highlightbackground": COLORS["blue_dark"],
            "highlightcolor": COLORS["blue_light"],
        }

        buttons = [
            ("NEW PROJECT", self.open_new_project),
            ("OPEN EXITING PROJECT",self.open_exiting_project),
            ("COPY EXITING PROJECT",self.open_exiting_project),
        ]

        for i, (text, command) in enumerate(buttons):
            button = tk.Button(self.root, text=text, command=command, **button_config)
            button.grid(row=2, column=i, pady=5, padx=5, sticky="nsew")

    def change_language(self):
        """Changes the language of the UI."""
        if self.language == "En.Language":
            self.language = "Per.Language"
            messagebox.showinfo("Language Changed", "Language switched to Persian.")
        else:
            self.language = "En.Language"
            messagebox.showinfo("Language Changed", "Language switched to English.")

    def open_new_project(self):
        """Opens the new Project window."""
        CreateProject(self.root)
        
    def open_exiting_project(self):
        print("no project exited")

    def open_ups_battery_ampere(self):
        """Opens the UPS Battery Ampere Calculation window."""
        UPSBatteryAmpere(self.root)

    def open_ups_support_time(self):
        """Opens the UPS Support Time Calculation window."""
        UPSBatterySupport(self.root)
        

    
# Run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = Application(root)
    root.mainloop()