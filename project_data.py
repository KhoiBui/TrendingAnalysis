from openpyxl import load_workbook


class TrendData(object):

    PROCESS_AREAS = [['PP', 0, 0, 0, 0, 0],
                     ['IPM', 0, 0, 0, 0, 0],
                     ['PMC', 0, 0, 0, 0, 0],
                     ['RSKM', 0, 0, 0, 0, 0],
                     ['REQM', 0, 0, 0, 0, 0],
                     ['RD', 0, 0, 0, 0, 0],
                     ['TS', 0, 0, 0, 0, 0],
                     ['PI', 0, 0, 0, 0, 0],
                     ['VER', 0, 0, 0, 0, 0],
                     ['VAL', 0, 0, 0, 0, 0],
                     ['CM', 0, 0, 0, 0, 0],
                     ['MA', 0, 0, 0, 0, 0],
                     ['PPQA', 0, 0, 0, 0, 0],
                     ['DAR', 0, 0, 0, 0, 0],
                     ['SAM', 0, 0, 0, 0, 0],
                     ['All PA\'s', 0, 0, 0, 0, 0]]

    def __init__(self, project_info, process_areas, workbook, worksheet):
        self.project_info = project_info
        self.process_areas = process_areas
        self.workbook = load_workbook(workbook)
        self.worksheet = self.workbook.get_sheet_by_name(worksheet)

    def write_process_data(self):
        row_offset = self.worksheet.max_row + 4
