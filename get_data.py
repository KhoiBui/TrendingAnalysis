""" Extract the findings table from final CAPA report. """

import sys
import re
from docx import Document
from openpyxl import load_workbook

class GetData(object):
    """ Read the final CAPA and extract info. """

    project_info = {}
    table_data = []
    data_read = []
    table = ''
    findings = ['Process Area', 'Goal', 'Practice', 'Description', 'Rating']

    def __init__(self, workbook, worksheet, document):
        self.workbook = load_workbook(workbook, read_only=False)
        self.worksheet = self.workbook.get_sheet_by_name(worksheet)
        self.document = Document(document)

    def process_document(self):
        """ Process the document. """
        self.read_doc()
        self.find_table()
        self.read_table_data(self.table)
        self.project_info.update({'Project Name':self.data_read[2]})
        self.project_info.update({'Lead(s)':self.data_read[3]})
        self.project_info.update({'Date Reported':self.data_read[4]})

    def find_table(self):
        """ Locate the detailed findings table. """
        tables = self.document.tables
        header = []
        for table in tables:
            header_row = table.rows[0]
            header[:] = []
            for cell in header_row.cells:
                for para in cell.paragraphs:
                    header.append(para.text.strip(' '))
            # check if elements in findings is also in header
            cond = len(header) == 5 and header[4] == 'Rating'
            if cond or [x for x in self.findings for y in header if x in y] == self.findings:
                self.table = table
                return

        # no table found
        print("Not able to find \"Detail of Findings\" table.")
        print("Possible that project does not have any findings.")
        sys.exit()

    def read_doc(self):
        """ Read document and put info into list. """
        for para in self.document.paragraphs:
            text = para.text
            # skip blank lines
            if text.strip():
                # remove duplicated spaces
                text = ' '.join(text.split())
                self.fill_project_info(text)
                self.data_read.append(text)

    def fill_project_info(self, line_read):
        """ Get general information about the project from doc. """
        line_read = line_read.split(':', 1)
        line_read[0] = re.sub('[- ]', '', line_read[0])
        key_name = line_read[0].lower()

        if 'sap' in key_name:
            self.project_info.update({'SAP ID':line_read[1].strip(' ')})
        elif 'golive' in key_name:
            self.project_info.update({'Go Live Date':line_read[1]})
        elif 'customer' in key_name:
            site_name = re.sub('State|Lottery', '', line_read[1])
            site_name = site_name.strip(' ')
            self.project_info.update({'Site':site_name})

    def read_table_data(self, table):
        """ Read info in specified table. """
        data = []
        index = 0
        for row in table.rows:
            data.append([])
            for cell in row.cells:
                text_data = ''
                for para in cell.paragraphs:
                    text_data += para.text.strip(' ')
                data[index].append(text_data)
            index += 1

        # don't need header row anymore
        self.table_data = data[1:]
        # trim end of lists
        self.table_data = [row[:5] for row in self.table_data]

    def get_table_data(self):
        """ Return data in table. """
        return self.table_data

    def get_project_info(self):
        """ Return information about the project. """
        return self.project_info

    def get_doc_data(self):
        """ Return information about the project. """
        return self.data_read

    def get_worksheet(self):
        """ Return name of the worksheet. """
        return self.worksheet

    def get_workbook(self):
        """ Return name of the workbook. """
        return self.workbook

    def get_document(self):
        """ Return name of the document. """
        return self.document
