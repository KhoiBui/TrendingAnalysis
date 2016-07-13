""" Extract the findings table from final CAPA report. """

from openpyxl import load_workbook

class ExtractTable(object):
    """ help. """

    def __init__(self, workbook_name, month):
        self.workbook_name = workbook_name
        self.month = month

    def open_workbook(self):
        """ help. """
        workbook = load_workbook(self.workbook_name)
        worksheet = workbook.get_sheet_by_name(self.month)

        return worksheet

    def find_table(self):
        """ help. """
        pass

    def read_doc(self):
        """ help. """
        pass

    def read_table_data(self):
        """ help. """
        pass
