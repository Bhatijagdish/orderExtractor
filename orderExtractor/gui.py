import sys
import tkinter as tk
from tkinter import messagebox

from main import snapshot_creation, static_product_cost, channel_sheet_creation

data = {}


def get_static_product_cost() -> None:
    global data
    data = static_product_cost()


def get_labour_charges():
    try:
        value = enter_value.get() if str(enter_value.get()).isdigit() else 0
        return int(value)
    except:
        messagebox.showerror("Invalid Value", "Pack & Labour Charges should be integer value")


def generate_channel_data():
    pack_labour_charges = get_labour_charges()
    if type(pack_labour_charges) != int:
        return None

    ad_expenses_checked = ad_checkbox.get()

    channel_sheet_creation(data, bool(ad_expenses_checked), pack_labour_charges)


root = tk.Tk()
root.title("Orders Report Generator")

root.geometry("500x200")

pack_and_labour_charges = tk.Label(root, text="Pack & Labour Charges")
pack_and_labour_charges.pack(padx=10, pady=10)

enter_value = tk.Entry(root)
enter_value.pack(padx=10, pady=10)

ad_checkbox = tk.IntVar()
ad_expenses = tk.Checkbutton(root, text="Ad Expenses", variable=ad_checkbox)
ad_expenses.pack(padx=10, pady=10)


select_orders_file_button = tk.Button(root, text="Prepare Sales Snapshot", command=snapshot_creation)
select_orders_file_button.pack(side=tk.LEFT, padx=20)

select_cost_data_file_button = tk.Button(root, text="Select SKU Cost Data", command=get_static_product_cost)
select_cost_data_file_button.pack(side=tk.LEFT, padx=20)


get_net_realisation_button = tk.Button(root, text="Get Net Realisation", command=generate_channel_data)
get_net_realisation_button.pack(side=tk.LEFT, padx=20)

root.resizable(False, False)

root.mainloop()
