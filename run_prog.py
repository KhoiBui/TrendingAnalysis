import docx_to_xlsx
import os, sys
import tkinter as tk
from tkinter import filedialog, Frame, BOTH, Button, RIGHT, RAISED, X,\
                    LEFT, Label, Entry, CENTER, StringVar

class TrendGUI(Frame):

    months = ['January', 'February', 'March', 'April', 'May', 'June',
              'July', 'August', 'September', 'October', 'November', 'December']

    def __init__(self, parent):
        Frame.__init__(self, parent, background='white')
        # saved reference to parent widget. "Tk root window"
        self.parent = parent
        self.month_var = StringVar()
        self.initGUI()

    def initGUI(self):
        """ Create the GUI. """
        # set title of root window
        self.parent.title('Trending Analysis Program')
        # fill frame to take up whole of root window
        self.pack(fill=BOTH, expand=True)

        # Fields
        frame = Frame(self, relief=RAISED)
        frame.pack(fill=BOTH, expand=True)
        month_lbl = Label(frame, text='Month', width=6)
        month_lbl.pack(side=LEFT, padx=5, pady=5)

        self.month_entry = Entry(frame, textvariable=self.month_var, justify=CENTER)
        self.month_entry.pack(fill=X, padx=5, expand=True)
        self.month_entry.bind('<Return>', self.get_month)

        # Buttons
        close_button = Button(self, text='Close', command=self.quit)
        close_button.pack(side=RIGHT, padx=5, pady=5)
        folder_button = Button(self, text='Select Folder')
        folder_button.pack(side=RIGHT)
        file_button = Button(self, text='Select File')
        file_button.pack(side=RIGHT, padx=5, pady=5)

    def get_month(self, *args):
        """ Return the input month. """
        value = self.month_entry.get()
        if value not in self.months:
            return None
        self.month_var = value

def main():

    root = tk.Tk()
    root.geometry("250x100+300+300")
    app = TrendGUI(root)
    root.mainloop()

if __name__ == '__main__':
    main()