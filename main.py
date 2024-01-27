import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import openpyxl
import subprocess

def check_combobox_state(event):
    v = status_combobox.get() 
    if v:
        save_btn.config(state=tk.NORMAL)
    else:
        save_btn.config(state=tk.DISABLED)

def exit_program():
    root.destroy()

def save():
    folder_path = filedialog.askdirectory()
    # print(folder_path)
    # print(type(folder_path))

    filtered_items = status_combobox.get()
    # print(filtered_items)
    # print(list_values)
    # print(type(list_values))
    # df = pd.DataFrame(list_values)
    # print(df.head)
    lists = [list(t) for t in list_values]
    # print(type(lists))
    df = pd.DataFrame(lists[1:],columns=lists[0])
    for i, j in enumerate(df[filtered_items].drop_duplicates()):
        df_filtered = df[df[filtered_items]==j]
        # print(df_filtered)
        # df_filtered.to_csv(f'{folder_path}\\{i+1}. {j}.csv', index=False)

    cols = [1,2,3,4]
    tree = ttk.Treeview(disp_frame, show="headings", columns=cols, height=12)
    tree.grid()

    status_combobox.set('')
    status_combobox.state(['disabled'])

    save_btn = ttk.Button(execute_frame, text="save", command=save, state=tk.DISABLED)
    save_btn.grid(row=0, column=0)

    directory_path = rf'{folder_path.replace('/','\\')}'
    subprocess.Popen(f'explorer "{directory_path}"')

def load_data(filepath):
    workbook = openpyxl.load_workbook(filepath)
    sheet = workbook.active
    global list_values
    list_values = list(sheet.values)
    # print(list_values[0][0])
    # lists = [list(t) for t in list_values]
    # print(lists[0:10])

    column_names = list_values[0]
    # mengupdate treeview
    global tree
    w = tree.winfo_width()
    tree.grid_forget()
    tree = ttk.Treeview(disp_frame, show="headings", columns=column_names, height=12)

    for col in column_names:
        tree.heading(col, text=col)
        # formatting
        tree.column(col, width=(int(w/6)) )

    for i in list_values[1:]:
        tree.insert('',tk.END,values=i)

    tree.place(relx=0, rely=0, width=w)

    global status_combobox
    status_combobox = ttk.Combobox(control_frame, values=column_names, state=tk.NORMAL)
    status_combobox.grid(row=0, column=1, padx=10, pady=5)
    status_combobox.bind("<<ComboboxSelected>>", check_combobox_state)


def select_file():
    file_path = filedialog.askopenfilename(title="Select File", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        # print(f"Selected file: {file_path}")
        load_data(file_path)

######################################################################################################
######################################################################################################
######################################################################################################
         
combo_list = ['']

root = tk.Tk()
# root.maxsize(width=800, height=600)
# root.geometry('800x600')

# MAIN FRAME
mainframe = ttk.Frame(root)
mainframe.pack(pady=(0,20))

# TITLE FRAME
# title_frame = ttk.Frame(mainframe)
# title_frame.grid(row=0,column=0,columnspan=2)

# title_label = ttk.Label(title_frame,text="Lorem Ipsum")
# title_label.pack(pady=10)

# FRAME MENU
menu_frame = ttk.LabelFrame(mainframe, text="Insert Excel File")
menu_frame.grid(row=0, column=0, padx=20, pady=10, sticky='w')

choose_file_btn = ttk.Button(menu_frame, text='Choose File', command=select_file)
choose_file_btn.grid(row=0, column=0, sticky='EW',padx=5, pady=5,)

# DISPLAY FRAME
disp_frame = ttk.Frame(mainframe)
disp_frame.grid(row=1,column=0,pady=10,padx=10)

cols = [1,2,3,4]
tree = ttk.Treeview(disp_frame, show="headings", columns=cols, height=12)
tree.grid()

# CONTROL FRAME
control_frame = ttk.Frame(mainframe)
control_frame.grid(row=2,column=0, sticky='W', padx=(10,0))

label1 = ttk.Label(control_frame, text='Pilih kolom untuk memfilter:')
label1.grid(row=0,column=0)

status_combobox = ttk.Combobox(control_frame, values=combo_list, state=tk.DISABLED)
status_combobox.grid(row=0, column=1, padx=10, pady=5)

execute_frame = ttk.Frame(mainframe)
execute_frame.grid(row=2,column=0, sticky='E',padx=(0,10))

save_btn = ttk.Button(execute_frame, text="save", command=save, state=tk.DISABLED)
save_btn.grid(row=0, column=0)

exit_btn = ttk.Button(execute_frame, text="exit", command=exit_program)
exit_btn.grid(row=0, column=1)

root.mainloop()