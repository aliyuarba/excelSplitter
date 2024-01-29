import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
import openpyxl
import subprocess

# FUNCTION TO EXECUTE WHEN TREEVIEW SCROLLING
def onxscroll(*args):
    tree_onload.xview(*args)
def onyscroll(*args):
    tree_onload.yview(*args)

# FUNCTION TO EXECUTE WHEN DROPDOWN MENU SELECTED
def comboFunction(event):
    print(dropdown_menu.get())
    save_btn.config(state=tk.NORMAL)

# FUNCTION TO CLEAR TREEVIEW DATA
def clear_data():
    # CLEARING DATA
    tree_onload.destroy()
    xscroll.destroy()
    yscroll.destroy()
    
    # DISABLING SAVE AND CLEAR BUTTON
    save_btn.config(state=tk.DISABLED)
    clear_btn.config(state=tk.DISABLED)

    # DISABLING DROPDOWN MENU AND CLEAR VALUE SELECTED
    dropdown_menu.set('')
    dropdown_menu.state(['disabled'])

    # ENABLING CHOOSE FILE BUTTON
    choose_file_btn.config(state=tk.NORMAL)

    global datadropped
    datadropped = False

# FUNCTION TO EXECUTE WHEN EXIT BUTTON CLICKED
def exit_program():
    root.destroy()

# LOAD DATA TO TREEVIEW
def load_data(filepath):
    # READ DATA FROM WORKBOOK USING OPENPYXL
    workbook = openpyxl.load_workbook(filepath)
    sheet = workbook.active

    # CONVERT WORKBOOK TO LIST
    global list_values
    list_values = list(sheet.values)

    # MAKE DATA VARIABLE FOR THE TREEVIEW HEADING
    heading_columns = list_values[0]
    
    # MAKE NEW TREEVIEW
    global tree_onload
    tree_onload = ttk.Treeview(disp_frame, show="headings", columns=heading_columns, height=14)
    
    global xscroll # MAKE HORIZONTAL SCROLLBAR
    xscroll = ttk.Scrollbar(disp_frame, orient="horizontal", command=onxscroll)
    
    global yscroll # MAKE VERTICAL SCROLLBAR
    yscroll = ttk.Scrollbar(disp_frame, orient="vertical", command=onyscroll)
    
    # CONFIGURE VERTICAL SCROLLBAR AND HORIZONTAL SCROLLBAR TO NEW TREEVIEW 
    tree_onload.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)

    # GET WIDTH AND HEIGHT OF INITIATED TREEVIEW
    w = tree.winfo_width()
    h = tree.winfo_height()

    for col in heading_columns:
        # SET HEADING COLUMN VARIABLE TO TREEVIEW HEADING COLUMN
        tree_onload.heading(col, text=col)
        # FORMATTING COLUMN
        tree_onload.column(col, width=(int(w/6)) )
    for i in list_values[1:]:
        # INSERT LIST VALUES TO TREEVIEW
        tree_onload.insert('',tk.END,values=i)

    # CALL TREEVIEW AND THE SCROLLBAR TO FRAME
    xscroll.place(x=0, y=h-20, height=20, width=w)
    yscroll.place(x=w-20, y=0, height=h, width=20)
    tree_onload.place(relx=0, rely=0, width=w)

    # DISABLING CHOOSE BUTTON
    choose_file_btn.config(state=tk.DISABLED)

    # ENABLING CLEAR BUTTON AND DROPDOWN MENU BUTTON
    dropdown_menu.config(values=heading_columns, state=tk.NORMAL)
    clear_btn.config(state=tk.NORMAL)

    global datadropped
    datadropped = True

# FUNCTION TO EXECUTE WHEN SAVE BUTTON CLICKED
def save():
    # OPEN DIALOG BOX TO SELECT DIRECTORY TO SAVE OUTPUT FILE 
    folder_path = filedialog.askdirectory()
    
    # GET VALUE FROM DROPDOWN MENU 
    filtered_items = dropdown_menu.get()

    # CONVERT TUPPLE LIST INTO ARRAY LIST
    lists = [list(t) for t in list_values]

    # CONVERT LIST INTO DATAFRAME
    df = pd.DataFrame(lists[1:],columns=lists[0])

    # ADD PROGRESS BAR INTO BOTTOM OF WINDOW
    progress_bar.grid(row=3,column=0,sticky='ew')

    # MAKE LIST FROM DATAFRAME BASED ON DROPDOWN MENU SELECTED BY USER
    item = df[filtered_items].drop_duplicates()
    
    for i, j in enumerate(item):
        # MAKE NEW DATAFRAME FROM MAIN DATAFRAME THAT FILTERED BY EACH ITEM
        df_filtered = df[df[filtered_items]==j]
        # CONVERT EACH NEW DATAFRAME TO MS EXCEL AND SAVE TO COMPUTER 
        # df_filtered.to_excel(f'{folder_path}\\{i+1}. {j}.xlsx', index=False)

        # UPDATING PROGRESSBAR
        progress_var.set((i+1)/len(item)*100)
        root.update_idletasks()
        progress_bar.update()

    # OPEN OUTPUT DIRECTORY FOLDER
    directory_path = rf'explorer {folder_path.replace('/','\\')}'
    # subprocess.Popen(f'explorer "{directory_path}"')
    # subprocess.Popen()
    confirm_box(directory_path)
    # DELETE PROGRESSBAR AFTER THE PROCESS DONE
    progress_bar.grid_forget()

def confirm_box(path):
    global custom_message_box
    custom_message_box = tk.Toplevel(root)

    frame1 = ttk.Frame(custom_message_box)
    frame1.grid(row=0,column=0)
    frame2 = ttk.Frame(custom_message_box)
    frame2.grid(row=1,column=0, sticky='e')

    label = ttk.Label(frame1, text="Splitted excel worksheet succesfully saved to your computer.\nDo you want to open the folder?", justify=tk.CENTER)
    label.grid(row=0,column=0, padx=20, pady=(20,30),sticky='ew',columnspan=2)

    # Yes button
    yes_button = ttk.Button(frame2, text="Yes", command=lambda :handle_yes(path))
    yes_button.grid(row=1,column=0, sticky='e', pady=(0,10))

    # No button
    no_button = ttk.Button(frame2, text="No", command=handle_no)
    no_button.grid(row=1,column=1, sticky='e',pady=(0,10),padx=(0,20))

def handle_yes(path):
    subprocess.Popen(path)
    custom_message_box.destroy()

def handle_no():
    custom_message_box.destroy()

def select_file(event):
    file_path = filedialog.askopenfilename(title="Select File", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        # print(f"Selected file: {file_path}")
        load_data(file_path)

def on_drop(event):
    if not datadropped:
        files = event.data.replace('{', '').replace('}','')
        load_data(files)
    else:
        # Show an alert if files have already been dropped
        messagebox.showinfo("Alert", "Files have already been dropped.")

# ROOT - MAIN WINDOW -
root = TkinterDnD.Tk()

# MAIN FRAME
mainframe = ttk.Frame(root)
mainframe.pack(pady=(0,20))

# FRAME MENU
menu_frame = ttk.LabelFrame(mainframe, text="Insert Excel File")
menu_frame.grid(row=0, column=0, padx=20, pady=10, sticky='w')

# CHOOSE FILE BUTTON
choose_file_btn = ttk.Button(menu_frame, text='Choose File', command=lambda: select_file('home'))
choose_file_btn.grid(row=0, column=0, sticky='EW',padx=5, pady=5,)

# DISPLAY FRAME
disp_frame = ttk.Frame(mainframe)
disp_frame.grid(row=1,column=0,pady=10,padx=10)

# INITIATE TREEVIEW
cols = [1,2,3,4]
tree = ttk.Treeview(disp_frame, show="headings", columns=cols, height=14)
tree.grid(row=0,column=0)
tree.bind("<Button-1>", select_file)

# DROP FILE PLACEHOLDER
label0 = ttk.Label(disp_frame, text='Just click this area \nor \ndrop your excel file here',background='#fff',foreground='gray', justify=tk.CENTER)
label0.grid(row=0,column=0)
label0.bind("<Button-1>", select_file)

# DROPDOWN FRAME
dropdown_frame = ttk.Frame(mainframe)
dropdown_frame.grid(row=2,column=0, sticky='W', padx=(10,0))

# SELECT COLUMN LABEL
dropdown_label = ttk.Label(dropdown_frame, text='Select Column:')
dropdown_label.grid(row=0,column=0)

# DROPDOWN MENU TO SELECT TABLE COLUMN
combo_list = ['']
dropdown_menu = ttk.Combobox(dropdown_frame, values=combo_list)
dropdown_menu.grid(row=0, column=1, padx=10, pady=5)
dropdown_menu.state(['disabled'])
dropdown_menu.bind("<<ComboboxSelected>>", comboFunction)

# SAVE EXIT FRAME CONTAINER
save_exit_frame = ttk.Frame(mainframe)
save_exit_frame.grid(row=2,column=0, sticky='E',padx=(0,10))

# SAVE BUTTON
save_btn = ttk.Button(save_exit_frame, text="save", command=save, state=tk.DISABLED)
save_btn.grid(row=0, column=0)

# CLEAR BUTTON
clear_btn = ttk.Button(save_exit_frame, text="clear", command=clear_data, state=tk.DISABLED)
clear_btn.grid(row=0, column=1)

# EXIT BUTTON
exit_btn = ttk.Button(save_exit_frame, text="exit", command=exit_program)
exit_btn.grid(row=0, column=2)

# MAKE PROGRESS BAR
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(mainframe, variable=progress_var, maximum=100)
progress_bar.grid_forget()

# Configure drop event
disp_frame.drop_target_register(DND_FILES)
disp_frame.dnd_bind('<<Drop>>', on_drop)
datadropped = False

root.mainloop()