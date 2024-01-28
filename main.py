import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import openpyxl
import subprocess

def onxscroll(*args):
    tree.xview(*args)
def onyscroll(*args):
    tree.yview(*args)

def comboFunction(event):
    print(combo.get())
    save_btn.config(state=tk.NORMAL)

def clear_treeview():
    # tree.destroy()
    cols = [1,2,3,4]
    tree = ttk.Treeview(disp_frame, show="headings", columns=cols, height=14)
    tree.grid(row=0,column=0)

    save_btn.config(state=tk.DISABLED)
    clear_btn.config(state=tk.DISABLED)

    combo.set('')
    combo.state(['disabled'])

def exit_program():
    root.destroy()

def save():
    folder_path = filedialog.askdirectory()

    filtered_items = combo.get()

    lists = [list(t) for t in list_values]
    # print(type(lists))

    df = pd.DataFrame(lists[1:],columns=lists[0])

    progress_bar.grid(row=3,column=0,sticky='ew')

    item = df[filtered_items].drop_duplicates()
    
    for i, j in enumerate(item):
        df_filtered = df[df[filtered_items]==j]
        # print(j)
        df_filtered.to_excel(f'{folder_path}\\{i+1}. {j}.xlsx', index=False)

        progress_var.set((i+1)/len(item)*100)
        root.update_idletasks()
        progress_bar.update()

    directory_path = rf'{folder_path.replace('/','\\')}'
    subprocess.Popen(f'explorer "{directory_path}"')

    progress_bar.grid_forget()

def load_data(filepath):
    workbook = openpyxl.load_workbook(filepath)
    sheet = workbook.active
    global list_values
    list_values = list(sheet.values)

    column_names = list_values[0]
    # mengupdate treeview
    global tree
    w = tree.winfo_width()
    h = tree.winfo_height()
    tree.grid_forget()
    tree = ttk.Treeview(disp_frame, show="headings", columns=column_names, height=14)

    xscroll = ttk.Scrollbar(disp_frame, orient="horizontal", command=onxscroll)
    yscroll = ttk.Scrollbar(disp_frame, orient="vertical", command=onyscroll)
    tree.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)

    for col in column_names:
        tree.heading(col, text=col)
        # formatting
        tree.column(col, width=(int(w/6)) )

    for i in list_values[1:]:
        tree.insert('',tk.END,values=i)

    xscroll.place(x=0, y=h-20, height=20, width=w)
    yscroll.place(x=w-20, y=0, height=h, width=20)
    tree.place(relx=0, rely=0, width=w)

    combo.config(values=column_names, state=tk.NORMAL)
    clear_btn.config(state=tk.NORMAL)

def select_file():
    file_path = filedialog.askopenfilename(title="Select File", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        # print(f"Selected file: {file_path}")
        load_data(file_path)

combo_list = ['']

root = tk.Tk()

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
tree = ttk.Treeview(disp_frame, show="headings", columns=cols, height=14)
tree.grid()

# CONTROL FRAME
control_frame = ttk.Frame(mainframe)
control_frame.grid(row=2,column=0, sticky='W', padx=(10,0))

label1 = ttk.Label(control_frame, text='Select Column:')
label1.grid(row=0,column=0)

combo = ttk.Combobox(control_frame, values=combo_list)
combo.grid(row=0, column=1, padx=10, pady=5)
combo.state(['disabled'])
combo.bind("<<ComboboxSelected>>", comboFunction)

execute_frame = ttk.Frame(mainframe)
execute_frame.grid(row=2,column=0, sticky='E',padx=(0,10))

save_btn = ttk.Button(execute_frame, text="save", command=save, state=tk.DISABLED)
save_btn.grid(row=0, column=0)

clear_btn = ttk.Button(execute_frame, text="clear", command=clear_treeview, state=tk.DISABLED)
clear_btn.grid(row=0, column=1)

exit_btn = ttk.Button(execute_frame, text="exit", command=exit_program)
exit_btn.grid(row=0, column=2)

progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(mainframe, variable=progress_var, maximum=100)
progress_bar.grid_forget()

root.mainloop()