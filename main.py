from tkinter import *
from myfuncs import *

window = Tk()
window.geometry("500x400+700+100")
window.title('Combine Excel')

for i in list(range(10)):
    window.columnconfigure(i, weight=1)

for i in list(range(10)):
    window.rowconfigure(i, weight=1)

# mengatur button 1
# button1 = Button(text="button1")
# button1.grid(row=9, column=0, sticky='WES')

# mengatur button 2
# button2 = Button(text="2")
# button2.grid(row=9, column=1, sticky='WES')

# mengatur button 3
button3 = Button(text="save",command=save)
button3.grid(row=9, column=6, sticky='WES',columnspan=2)
button3.configure(bg='#65B741',fg='#fff')

# mengatur button 4
button4 = Button(text="cancel",command=closewindow)
button4.grid(row=9, column=8, sticky='WES',columnspan=2)
button4.configure(bg='#EF4040',fg='#fff')

window.mainloop()