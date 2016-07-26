""" Extract data from a specific table in Final CAPA reports
    and put it into trending analysis spreadsheet. """

import sys
import get_data
import write_data


def main(argv, workbook_name, worksheet_name):
    """ Run the program. """
    print('Loading {}'.format(argv))
    project = get_data.GetData(argv, workbook_name, worksheet_name)
    project.process_document()
    workbook = project.get_workbook()
    worksheet = project.get_worksheet()
    table_data = project.get_table_data()
    project_info = project.get_project_info()

    # write to spreadsheet
    do_write = write_data.WriteData(worksheet, table_data, project_info)
    do_write.write_to_sheet()
    project_info = do_write.get_project_info()
    process_areas = do_write.get_process_areas()

    # save changes
    print('Saving to {}'.format(workbook_name))
    workbook.save(workbook_name)
    print('Done.')

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3])
