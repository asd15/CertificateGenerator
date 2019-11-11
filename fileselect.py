from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import auto

class Root(Tk):
    def __init__(self):
        super(Root, self).__init__()
        self.title("Certificate Automation")
        self.minsize(640, 400)

        self.labelFrame = ttk.LabelFrame(self, text="Import a File")
        self.labelFrame.grid(column=0, row=1, padx=50, pady=50)

        self.button()

    def button(self):
        self.button = ttk.Button(self.labelFrame, text="Select an EXCEL file to send Ceritificates", command=self.fileDialog)
        self.button.grid(column=1, row=1)

    def fileDialog(self):
        self.filename = filedialog.askopenfilename(initialdir="/", title="Select a file", filetype=(("xlsx", "*.xlsx"), ("All Files", "*.*")))
        self.label = ttk.Label(self.labelFrame, text="")
        self.label.grid(column=1, row=2)
        self.label.configure(text=self.filename)
        self.label.configure(text="Emails Sent!")
        filename = self.filename
        print(filename)
        auto.takefile(filename)

if __name__ == '__main__':
    root = Root()
    root.mainloop()