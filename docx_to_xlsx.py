""" Extract data from a specific table in Final CAPA reports
    and put it into trending analysis spreadsheet. """

import sys
import get_data
import write_data


def main(argv):
    """ Run the program. """
    print('Loading {}'.format(argv))
    destination = 'Draft_Detail_Findings.xlsx'
    project = get_data.GetData(argv, destination)
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
    print(process_areas)
    print(project_info)

    # save changes
    print('Saving to {}'.format(destination))
    workbook.save(destination)
    print('Done.')

if __name__ == '__main__':
    main(sys.argv[1])
