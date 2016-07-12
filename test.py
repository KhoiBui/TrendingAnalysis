import sys
import re
from openpyxl import *
from openpyxl.styles import *
from docx import Document

""" This dictionary was providing inconsistent ordering of the key
    and value pairs with every run. I thought this was a weird
    concurrency issue, but I never used multithreading in this
    program so there was no way for the global dictionary to be
    used by multiple threads at once. Turns out Python salts the 
    keys and shuffles the order for security reasons. 
    more info @ __hash__ """

project_info = {}

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

    wb = load_workbook('TestSheet.xlsx', read_only=False, )
    ws = wb.get_sheet_by_name('June Data')

    try:
        doc = Document("Example_Doc.docx")
    except OSError:
        print('Could not open the document, check that the file name is correct')
        sys.exit()

    findings = ['Process Area', 'Goal', 'Practice', 'Description', 'Rating']

    # read and process data in document
    info = readDoc(doc)
    # find the detail findings table in list of tables
    findings_table = findTable(doc, findings)
    # put data in found table into a list
    table_data = readTableData(findings_table)

    # fill out project information
    project_info.update({'Batch':info[2]})
    project_info.update({'Lead(s)':info[3]})
    project_info.update({'Report Date':info[4]})

    print(project_info)
    print()
    print(table_data)
    print()

    # trying to find first empty cell index
    col_index = 0
    row_index = ws.max_row + 1

    for row in ws.iter_rows():
        for cell in row:
            if cell.value == 'Process Area':
                col_index = cell.column
                break
            if cell.value is None:
                row_index = cell.row + 1
                break

    print("{}{}".format(col_index, row_index))
    print()

    style_center = Style(alignment=Alignment(horizontal='general',
                                             vertical='center',
                                             wrap_text=True))

    for row in range(len(table_data) - 1):
        for col in range(ord(col_index), ord(col_index) + 5):
            new_row = row + row_index
            new_col = col - 64
            
            working_cell = ws.cell(row=new_row, column=new_col)
            working_cell.style = style_center
            working_cell.value = str(table_data[row + 1][new_col - 5])

    wb.save('TestSheet.xlsx')

def findTable(doc, row_to_find):
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

def readTableData(table):
    """ Put data in table into a list. """
    data = []
    index = -1
    for row in table.rows:
        data.append([])
        index += 1
        for cell in row.cells:
            for para in cell.paragraphs:
                data[index].append(para.text.strip(' '))

    return data

def readDoc(doc):
    """ Read the document and look for specific information. """
    data_read = []
    for para in doc.paragraphs:
        text = para.text
        # skip blank lines
        if text == '':
            continue
        # remove duplicated spaces
        text = ' '.join(text.split())
        fillProjectInfo(text)
        data_read.append(text)

    return data_read

def fillProjectInfo(line_read):
    """ Find the rest of the project's information. """
    line_read = line_read.split(':', 1)
    line_read[0] = re.sub('[- ]', '', line_read[0])
    key_name  = line_read[0].lower()
    if 'sap' in key_name:
        project_info.update({'SAP ID':line_read[1]})
    elif 'golive' in key_name:
        project_info.update({'Go Live Date':line_read[1]})


if __name__ == '__main__':
    main()

