import tkinter as tk
from tkinter import ttk, messagebox
from utilities import COLORS

class UPSBatteryAmpere:
    """Class to calculate UPS Battery Ampere."""

    def __init__(self, root):
        self.root = root
        self.create_window()

    def create_window(self):
        """Creates the UPS Battery Ampere Calculation window."""
        self.window = tk.Toplevel(self.root)
        self.window.title("UPS Battery Ampere Calculation")
        self.window.geometry("400x400")
        self.window.configure(bg=COLORS["cream"])

        self.create_label_entry("Load (KW):", 0)
        self.create_label_entry("Support Time (min):", 1)
        self.create_label_entry("Battery Quantity:", 2)
        self.create_label_combo("Battery Efficiency (%):", 3, values=(40, 50, 60, 70, 80, 90, 100))

        # Output Labels
        self.output_labels = {
            "Load (KVA)": self.create_output_label("Load (KVA):", 4),
            "Battery Voltage (V)": self.create_output_label("Battery Voltage (V):", 5),
            "Battery Ampere (A)": self.create_output_label("Battery Ampere (A):", 6)
        }

        calc_button = tk.Button(
            self.window, text="Calculate", bg=COLORS["green_fosfori"],
            command=self.calculate_battery_ampere
        )
        calc_button.grid(row=7, column=0, columnspan=2, pady=10)

    # Methods for creating UI components
    def create_label_entry(self, text, row):
        """Creates a label and entry field."""
        label = ttk.Label(self.window, text=text, background=COLORS["cream"], width=22)
        label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
        var = tk.DoubleVar()
        entry = ttk.Entry(self.window, textvariable=var, width=12)
        entry.grid(row=row, column=1, padx=5, pady=5)
        setattr(self, f"input_{text.split()[0].lower()}", var)

    def create_label_combo(self, text, row, values):
        """Creates a label and combobox field."""
        label = ttk.Label(self.window, text=text, background="white")
        
        
class UPSBatterySupport:
    """Class to calculate UPS Battery Ampere."""

    def __init__(self, root):
        self.root = root
        self.create_window()

    def create_window(self):
        """Creates the UPS Battery Ampere Calculation window."""
        self.window = tk.Toplevel(self.root)
        self.window.title("UPS Battery Ampere Calculation")
        self.window.geometry("400x400")
        self.window.configure(bg=COLORS["cream"])

        self.create_label_entry("Load (KW):", 0)
        self.create_label_entry("Support Time (min):", 1)
        self.create_label_entry("Battery Quantity:", 2)
        self.create_label_combo("Battery Efficiency (%):", 3, values=(40, 50, 60, 70, 80, 90, 100))

        # Output Labels
        self.output_labels = {
            "Load (KVA)": self.create_output_label("Load (KVA):", 4),
            "Battery Voltage (V)": self.create_output_label("Battery Voltage (V):", 5),
            "Battery Ampere (A)": self.create_output_label("Battery Ampere (A):", 6)
        }

        calc_button = tk.Button(
            self.window, text="Calculate", bg=COLORS["green_fosfori"],
            command=self.calculate_battery_ampere
        )
        calc_button.grid(row=7, column=0, columnspan=2, pady=10)

    # Methods for creating UI components
    def create_label_entry(self, text, row):
        """Creates a label and entry field."""
        label = ttk.Label(self.window, text=text, background=COLORS["cream"], width=22)
        label.grid(row=row, column=0, padx=5, pady=5, sticky="w")
        var = tk.DoubleVar()
        entry = ttk.Entry(self.window, textvariable=var, width=12)
        entry.grid(row=row, column=1, padx=5, pady=5)
        setattr(self, f"input_{text.split()[0].lower()}", var)

    def create_label_combo(self, text, row, values):
        """Creates a label and combobox field."""
        label = ttk.Label(self.window, text=text, background="white")