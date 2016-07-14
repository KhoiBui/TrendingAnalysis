""" test. """

import get_data

def main():
    """ test. """
    proc = get_data.GetData('TestSheet.xlsx', 'June Data', 'Example_Doc.docx')
    proc.process_document()
    print(proc.get_table_data())
    print()
    print(proc.get_project_info())
    print()
    print(proc.get_doc_data())

if __name__ == '__main__':
    main()
