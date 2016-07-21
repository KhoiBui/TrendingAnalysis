""" Extract data from a specific table in Final CAPA reports
    and put it into trending analysis spreadsheet. """

import sys
import get_data
import write_data

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
        - Process Area   Goal   Practice   Description   Rating """


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

    # save changes
    print('Saving to {}'.format(destination))
    workbook.save(destination)

if __name__ == '__main__':
    main(sys.argv[1])
