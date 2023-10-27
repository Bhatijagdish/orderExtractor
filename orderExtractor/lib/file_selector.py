import sys
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


def select_file():
    root = tk.Tk()
    try:
        root.withdraw()

        file_path = filedialog.askopenfilename(
            filetypes=[('CSV & Excel Files', "*.csv*;*.xls*;*.xlsx*;*.xl*;")]
        )
        if file_path:
            return file_path
        else:
            messagebox.showinfo("File Selection", "Please select at least one file that contains orders data")
            sys.exit(0)
    except:
        messagebox.showerror("Error", "Unable to find an orders file path")
        sys.exit(0)
