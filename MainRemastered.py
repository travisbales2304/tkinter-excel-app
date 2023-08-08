import tkinter as tk
from tkinter import ttk
import openpyxl


def load_data():
    path = "C:\\Users\\travi\\OneDrive\\Documents\\GitHub\\tkinter-excel-app\\people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)
    print(list_values)
    for col_name in list_values[0]:
        treeview.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)

def insert_row():
    name = name_entry.get()
    age = int(age_spinbox.get())
    subscription_status = status_combobox.get()
    employment_status = "Employed" if a.get() else "Unemployed"

    print(name, age, subscription_status, employment_status)

    # Insert row into Excel sheet
    path = "C:\\Users\\travi\\OneDrive\\Documents\\GitHub\\tkinter-excel-app\\people.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [name, age, subscription_status, employment_status]
    sheet.append(row_values)
    workbook.save(path)

    # Insert row into treeview
    treeview.insert('', tk.END, values=row_values)
    
    # Clear the values
    name_entry.delete(0, "end")
    name_entry.insert(0, "CustomerName")
    age_spinbox.delete(0, "end")
    age_spinbox.insert(0, "SerialNumber")
    status_combobox.set(combo_list[0])
    checkbutton.state(["!selected"])

def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")
    treeview.column("CustomerName", width=100)
    treeview.column("SerialNumber", width=50)
    treeview.column("IsQuoted", width=100)
    treeview.column("ItemNumber", width=100)

root = tk.Tk()
root.title("Renewal Tracker")



icon_path = "C:\\Users\\travi\\OneDrive\\Documents\\GitHub\\tkinter-excel-app\\file_accept_checklist_check_document_icon_251831.ico"
root.iconbitmap(icon_path)

style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

combo_list = ["Approved", "Quote Sent", "Needs Quote"]

frame = ttk.Frame(root)
frame.pack()

tabControl = ttk.Notebook(root)
tab1 = ttk.Frame(tabControl)
tab2 = ttk.Frame(tabControl)
tabControl.add(tab1, text="Tab 1")
tabControl.add(tab2, text="Tab 2")
tabControl.pack(expand=1,fill="both")

widgets_frame = ttk.LabelFrame(tab2, text="Add Entry")
widgets_frame.grid(row=0, column=0, padx=20, pady=10)

name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "CustomerName")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))
name_entry.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew")

age_spinbox = ttk.Entry(widgets_frame)
age_spinbox.insert(0, "SerialNumber")
age_spinbox.bind("<FocusIn>", lambda e: age_spinbox.delete('0', 'end'))
age_spinbox.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

status_combobox = ttk.Combobox(widgets_frame, values=combo_list)
status_combobox.current(2)
status_combobox.grid(row=2, column=0, padx=5, pady=5,  sticky="ew")

a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widgets_frame, text="ItemNumber", variable=a)
checkbutton.grid(row=3, column=0, padx=5, pady=5, sticky="nsew")

button = ttk.Button(widgets_frame, text="Add", command=insert_row)
button.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame)
separator.grid(row=5, column=0, padx=(20, 10), pady=10, sticky="ew")

mode_switch = ttk.Checkbutton(widgets_frame, text="Mode", style="Switch", command=toggle_mode)
mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky="nsew")


treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

cols = ("CustomerName","SerialNumber", "IsQuoted", "ItemNumber")
treeview = ttk.Treeview(treeFrame, show="headings",
                        yscrollcommand=treeScroll.set, columns=cols, height=13)
treeview.column("CustomerName", width=100)
treeview.column("SerialNumber", width=50)
treeview.column("IsQuoted", width=100)
treeview.column("ItemNumber", width=100)


treeview.pack()
treeScroll.config(command=treeview.yview)
load_data()


root.mainloop()
