""" Extract data from a specific table in Final CAPA reports
    and put it into trending analysis spreadsheet. """

import sys
import re
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill
from docx import Document

""" This dictionary was providing inconsistent ordering of the key
    and value pairs with every run. I thought this was a weird
    concurrency issue, but I never used multithreading in this
    program so there was no way for the global dictionary to be
    used by multiple threads at once. Turns out Python salts the
    keys and shuffles the order for security reasons.
    more info @ __hash__ """

PROJECT_INFO = {}

""" Some values will have to be hardcoded in. This is because there
    are some inconsistencies in the Final CAPA's.

    Inconsistent:
    - Two leads for one project, separated by "/" E.g. "Adam/Monika"
    - Batch name sometimes not included under "Project Information"
    - Customer name is sometimes the site name, other times it is
      site name + entity
    - Different formatting/spelling of go live date E.g. "GO date" or
      "Go live date" or "Go-Live date"
    - Detail of Findings table is sometimes 2nd from last table or 3rd
      from last
        - some reports have extra color code table at the end

    Consistent:
    - Batch name is always 3rd non-empty line in document
    - followed by lead's name
    - followed by report date
    - SAP ID
    - "Project Stakeholders" section follows the same order
    - Detail of Findings table have consistent header
        - Process Area   Goal   Practice   Descrition   Rating """

def main():
    """ Run the program. """

    workbook = load_workbook('TestSheet.xlsx', read_only=False, )
    worksheet = workbook.get_sheet_by_name('June Data')

    try:
        doc = Document("Example_Doc.docx")
    except OSError:
        print('Could not open the document, check that the file name is correct')
        sys.exit()

    findings = ['Process Area', 'Goal', 'Practice', 'Description', 'Rating']

    # read and process data in document
    info = read_doc(doc)
    # find the detail findings table in list of tables
    findings_table = find_table(doc, findings)
    # put data in found table into a list
    table_data = read_table_data(findings_table)

    # fill out project information
    PROJECT_INFO.update({'Project Name':info[2]})
    PROJECT_INFO.update({'Lead(s)':info[3]})
    PROJECT_INFO.update({'Date Reported':info[4]})

    print(PROJECT_INFO)
    print()
    print(table_data)
    print()

    # trying to find first empty cell index
    col_offset = 0
    row_offset = 0

    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value == 'Process Area':
                col_offset = cell.column
                break
            if cell.value is None:
                row_offset = cell.row + 1
                break

    print("{}{}".format(col_offset, row_offset))
    print()

    # put data into the excel worksheet
    for row in range(len(table_data) - 1):
        # data from "findings" table
        fill_in_table_data(worksheet, table_data, row, row_offset, ord(col_offset))
        fill_in_project_info(worksheet, row, row_offset)
    # save changes made
    workbook.save('TestSheet.xlsx')

def fill_in_table_data(worksheet, table_data, row, row_offset, col_offset):
    """ Helper. """

    """ #92D050 - green (LI)
        #FFC000 - yellow (PI)
        #FF0000 - red (NI) """
    for col in range(col_offset, col_offset + 5):
        new_row = row + row_offset
        new_col = col - 64
        header_info = worksheet.cell(row=1, column=new_col).value
        working_cell = worksheet.cell(row=new_row, column=new_col)
        align = 'center'
        # put data into cell
        working_cell.value = str(table_data[row + 1][new_col - 5])

        if header_info == 'Finding':
            align = 'general'

        if header_info == 'Rating':
            color = ''
            if working_cell.value == 'LI':
                color = '92D050'
            elif working_cell.value == 'PI':
                color = 'FFC000'
            elif working_cell.value == 'NI':
                color = 'FF0000'
            working_cell.fill = PatternFill(fill_type='solid',
                                            start_color=color)

        working_cell.alignment = Alignment(horizontal=align,
                                           vertical='center',
                                           wrap_text=True)

def fill_in_project_info(worksheet, row, row_offset):
    """ Helper. """
    for col2 in range(1, 5, 1):
        working_cell = worksheet.cell(row=row_offset + row, column=col2)
        working_cell.alignment = Alignment(horizontal='center',
                                           vertical='center',
                                           wrap_text=True)
        header_info = worksheet.cell(row=1, column=col2).value.strip(' ')

        if header_info == 'Project Name':
            working_cell.value = PROJECT_INFO['Project Name']
        elif header_info == 'SAP ID':
            working_cell.value = PROJECT_INFO['SAP ID']
        elif header_info == 'Date Reported':
            working_cell.value = PROJECT_INFO['Date Reported']

def find_table(doc, row_to_find):
    """ Look for the table that contains the detailed findings. """
    if row_to_find is None:
        raise ValueError('Row header is invalid')

    tables = doc.tables
    header = []
    for table in tables:
        header_row = table.rows[0]
        header[:] = []      # could be slow, do benchmark later
        for cell in header_row.cells:
            for paragraph in cell.paragraphs:
                header.append(paragraph.text.strip(' '))
        if header == row_to_find:
            return table

def read_table_data(table):
    """ Put data in table into a list. """
    if table is None:
        raise ValueError('Table can\'t be null')

    data = []
    index = -1
    for row in table.rows:
        data.append([])
        index += 1
        for cell in row.cells:
            for para in cell.paragraphs:
                data[index].append(para.text.strip(' '))

    return data

def read_doc(doc):
    """ Read the document and look for specific information. """
    data_read = []
    for para in doc.paragraphs:
        text = para.text
        # skip blank lines
        if text == '':
            continue
        # remove duplicated spaces
        text = ' '.join(text.split())
        fill_project_info(text)
        data_read.append(text)

    return data_read

def fill_project_info(line_read):
    """ Find the rest of the project's information. """
    line_read = line_read.split(':', 1)
    line_read[0] = re.sub('[- ]', '', line_read[0])
    key_name = line_read[0].lower()
    if 'sap' in key_name:
        PROJECT_INFO.update({'SAP ID':line_read[1]})
    elif 'golive' in key_name:
        PROJECT_INFO.update({'Go Live Date':line_read[1]})

if __name__ == '__main__':
    main()
