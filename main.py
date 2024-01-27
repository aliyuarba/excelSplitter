from tkinter import *

window = Tk()
window.geometry("500x400+700+100")
window.title('Combine Excel')

for i in list(range(10)):
    window.columnconfigure(i, weight=1)

for i in list(range(10)):
    window.rowconfigure(i, weight=1)

# mengatur button 1
button1 = Button(text="button1")
button1.grid(row=9, column=0, sticky='WES')

# mengatur button 2
button2 = Button(text="2")
button2.grid(row=9, column=1, sticky='WES')

# mengatur button 3
button3 = Button(text="3")
button3.grid(row=9, column=2, sticky='WES',columnspan=7)

# mengatur button 4
button4 = Button(text="4")
button4.grid(row=9, column=9, sticky='WES')

window.mainloop()