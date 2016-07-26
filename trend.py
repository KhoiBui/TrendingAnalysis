""" Run the program with a GUI. """

import docx_to_xlsx
import os
import subprocess
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
            # check that file(s) selected is .docx NOT .doc
            self.check_ext()
            self.file_button.pack(fill=BOTH, expand=True, padx=5, pady=5)
            self.folder_button.destroy()

    def get_folder(self):
        self._folder_path = filedialog.askdirectory()
        if self._folder_path != '':
            self.folder_button.config(text='Folder Selected!')
            # check that file(s) selected is .docx NOT .doc
            self.check_ext()
            self.folder_button.pack(fill=BOTH, expand=True, padx=5, pady=5)
            self.file_button.destroy()

    def run_program(self):
        workbook = 'Draft_Detail_Findings.xlsx'
        worksheet = 'Template'
        # user selected one CAPA
        if self._folder_path is None:
            docx_to_xlsx.main(self._file_path, workbook, worksheet)
        # user selected a folder of CAPA's
        elif self._file_path is None:
            for f in os.listdir(self._folder_path):
                # get full path name
                file_name = str(self._folder_path + '/' + f)
                docx_to_xlsx.main(file_name, workbook, worksheet)

        # get ready to end the program
        self.frame_1.destroy()
        self.run_button.destroy()
        self.close_button.config(text='Done.')
        self.close_button.pack(fill=BOTH, expand=True, padx=10, pady=10)

    def check_ext(self):
        if self._file_path is not None:
            self._file_path = self.convert_to_docx(self._file_path)
            print('Ready. Click \'Run\' to Proceed.')
        elif self._folder_path is not None:
            for f in os.listdir(self._folder_path):
                file_name = str(self._folder_path + '/' + f)
                self.convert_to_docx(file_name)
            print('Ready. Click \'Run\' to Proceed.')
        else:
            raise OSError('File(s) does not exist or you did not select anything. ')

    @classmethod
    def convert_to_docx(cls, user_input):
        if str(user_input).endswith('.docx'):
            return user_input
        else:
            new_file_name = user_input.split(' ')
            new_file_name = new_file_name[0] + '_' + new_file_name[1] + '_CAPA.docx'
            word_conv = r'C:\Program Files (x86)\Microsoft Office\Office12\wordconv.exe'
            commands = ['wordconv.exe', '-oice', '-nme', user_input, new_file_name]
            try:
                print('Converting {} to {}'.format(user_input, new_file_name))
                subprocess.Popen(commands, executable=word_conv)
                # wait for converted file to be created
                while not os.path.exists(new_file_name):
                    time.sleep(1)
                print('Removing old .doc file ...')
                os.remove(user_input)
                return new_file_name
            except OSError:
                print('Failed to convert file(s). Check to see if it exists.')


def main():
    """ Run the gui and program. """
    root = tk.Tk()
    root.geometry("250x100+300+300")
    TrendProg(root)
    root.mainloop()

if __name__ == '__main__':
    main()
