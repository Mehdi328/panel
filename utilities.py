#===================================== Utilities =====================================

# Define global colors
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

def create_space(frame, x, y, w, bg):
    """Creates an empty space in the grid layout."""
    from tkinter import ttk
    ttk.Label(frame, text=" ", width=w, background=bg).grid(
        row=x, column=y, rowspan=1, columnspan=1, padx=3, pady=3, sticky='SNEW'
    )