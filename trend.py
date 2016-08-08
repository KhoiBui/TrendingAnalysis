""" Run the program with a GUI. """

import docx_to_xlsx
import os
import subprocess
import re
import time
import tkinter as tk
from tkinter import filedialog, Frame, BOTH, Button, RIGHT, RAISED,\
                    LEFT


class TrendProg(Frame):

    def __init__(self, parent):
        Frame.__init__(self, parent, background='white')
        # saved reference to parent widget. "Tk root window"
        self.parent = parent
        self._workbook = None
        self._file_path = None
        self._folder_path = None

        self.frame_1 = Frame(self, relief=RAISED)
        self.run_button = Button(self, text='Run', width=10,
                                 command=self.run_program)
        self.file_button = Button(self.frame_1, text='Select File',
                                  width=15, command=self.get_file)
        self.folder_button = Button(self.frame_1, text='Select Folder',
                                    width=15, command=self.get_folder)
        self.close_button = Button(self, text='Close', width=10,
                                   command=self.quit)
        self.init_gui()

    def init_gui(self):
        """ Create the GUI. """
        # set title of root window
        self.parent.title('Trending Analysis')
        # fill frame to take up whole of root window
        self.pack(fill=BOTH, expand=True)
        self.frame_1.pack(fill=BOTH, expand=True)

        # put buttons on GUI
        self.folder_button.pack(side=RIGHT, padx=5)
        self.file_button.pack(side=LEFT, padx=5, pady=5)
        self.close_button.pack(side=RIGHT, padx=5, pady=5)
        self.run_button.pack(side=RIGHT, pady=5)

    def get_file(self):
        self._file_path = filedialog.askopenfilename()
        if self._file_path != '':
            self.file_button.config(text='File Selected!')
            self.file_button.pack(fill=BOTH, expand=True, padx=5, pady=5)
            self.folder_button.destroy()

    def get_folder(self):
        self._folder_path = filedialog.askdirectory()
        if self._folder_path != '':
            self.folder_button.config(text='Folder Selected!')
            self.folder_button.pack(fill=BOTH, expand=True, padx=5, pady=5)
            self.file_button.destroy()

    def run_program(self):
        workbook = 'Draft_Detail_Findings.xlsx'
        worksheet = 'Template'
        # user selected one CAPA
        print('=' * 75)
        if self._folder_path == '' or self._folder_path is None:
            self._file_path = self.convert_to_docx(self._file_path)
            docx_to_xlsx.main(self._file_path, workbook, worksheet)
            print('=' * 75)
        # user selected a folder of CAPA's
        elif self._file_path == '' or self._file_path is None:
            for f in os.listdir(self._folder_path):
                # get full path name
                file_name = str(self._folder_path + '/' + f)
                file_name = self.convert_to_docx(file_name)
                docx_to_xlsx.main(file_name, workbook, worksheet)
                print('=' * 75)

        # get ready to end the program
        # pd = project_data.TrendData(workbook, worksheet)
        print('Done.')
        self.frame_1.destroy()
        self.run_button.destroy()
        self.close_button.config(text='Done.')
        self.close_button.pack(fill=BOTH, expand=True, padx=10, pady=10)

    @classmethod
    def convert_to_docx(cls, file_selected):
        """ Check that file(s) selected is .docx NOT .doc and convert if needed. """
        if str(file_selected).endswith('.docx'):
            return file_selected
        else:
            new_file_name = re.sub('.doc', '.docx', file_selected)
            # full path to wordconv.exe
            word_conv = r'C:\Program Files (x86)\Microsoft Office\Office12\wordconv.exe'
            commands = ['wordconv.exe', '-oice', '-nme', file_selected, new_file_name]
            try:
                print('CONVERTING {}'.format(file_selected))
                subprocess.Popen(commands, executable=word_conv)
                # wait for converted file to be created
                while not os.path.exists(new_file_name):
                    time.sleep(1.5)
                print('REMOVING old .doc file ...')
                os.remove(file_selected)
                return new_file_name
            except OSError:
                print('FAILED to convert file. Check to see if it exists.')


def main():
    """ Run the gui and program. """
    root = tk.Tk()
    root.geometry("250x100+300+300")
    TrendProg(root)
    root.mainloop()

if __name__ == '__main__':
    main()
