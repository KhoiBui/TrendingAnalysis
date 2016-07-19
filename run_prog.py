import docx_to_xlsx
import os, sys
import tkinter as tk
from tkinter import filedialog, Frame, BOTH, Button, RIGHT, RAISED

class TrendGUI(Frame):

    def __init__(self, parent):
        Frame.__init__(self, parent, background='white')
        # saved reference to parent widget. "Tk root window"
        self.parent = parent
        self.initGUI()

    def initGUI(self):
        """ Create the GUI. """
        # set title of root window
        self.parent.title('Trending Analysis Program')
        # fill frame to take up whole of root window
        self.pack(fill=BOTH, expand=True)

        frame = Frame(self, relief=RAISED, borderwidth=1)
        frame.pack(fill=BOTH, expand=True)
        close_button = Button(self, text='Close', command=self.quit)
        close_button.pack(side=RIGHT, padx=7, pady=7)
        folder_button = Button(self, text='Select Folder')
        folder_button.pack(side=RIGHT)

def main():

    root = tk.Tk()
    root.geometry("350x200+300+300")
    app = TrendGUI(root)
    root.mainloop()

if __name__ == '__main__':
    main()