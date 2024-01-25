from tkinter import *

window = Tk()
window.geometry("500x400+700+100")
window.title('Combine Excel')

window.columnconfigure(0, weight=1)
window.columnconfigure(1, weight=0)
window.columnconfigure(2, weight=1)
window.columnconfigure(3, weight=1)

window.rowconfigure(0, weight=1)
window.rowconfigure(1, weight=1)
window.rowconfigure(2, weight=1)
window.rowconfigure(3, weight=1)

button1 = Button(text="1")
button2 = Button(text="2")
button3 = Button(text="3")
button4 = Button(text="4")

button1.grid(row=1, column=3)
button2.grid(row=2, column=2)
button3.grid(row=3, column=4)
button4.grid(row=0, column=1)


window.mainloop()
