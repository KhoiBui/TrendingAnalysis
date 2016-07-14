""" test. """

import extract_table

def main():
    """ test. """
    proc = extract_table.ExtractTable('TestSheet.xlsx', 'June Data', 'Example_Doc.docx')
    proc.process_document()
    print(proc.get_table_data())
    print(proc.get_project_info())
    print(proc.get_doc_data())

if __name__ == '__main__':
    main()
