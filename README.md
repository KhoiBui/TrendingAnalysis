# TrendingProg
The purpose of this program is to assist with doing monthly trending analysis. The Standards Compliance team at IGT compiles a
spreadsheet of data from projects each month. This allows managers to observe trends as they develop and help to better
identify problematic process areas and their causes. Getting the data for trending analysis is a very manual task that involves
extracting certain pieces of information in Final CAPA Reports and putting them into a spreadsheet. This program attempts to
automate as much of that process as possible.


## Using the Program
The program reads a Word document and keeps track of information about the project as well as the "Detail of Findings" table. It
will then put this information into an Excel worksheet.

* Final CAPA's will have to be manually selected from the network drive and put into a folder. This is because leads will often
  save several different versions of a CAPA report and their names will vary.

  Draft_Detail_Findings.xlsx should be in the same working directory as the .exe file

  The directory structure should look something like this:
  
  ```
  TrendingProg
  ├── docx_to_xlsx.py
  ├── get_data.py
  ├── project_data.py
  ├── README.md
  ├── trend
  │   ├── CAPAs
  │   │   ├── April
  │   │   │   ├── CY1504_CAPA.docx
  │   │   │   ├── MI_CAPA.docx
  │   │   ├── July
  │   │   │   ├── CY1604 ES_Final CAPA Report.docx
  │   │   │   ├── CZH CY16_THE FINAL CAPA Report_20160727.doc
  │   │   │   ├── IN CY15_Final CAPA Report.doc
  │   │   │   ├── LUX CY16_THE FINAL CAPA Report_20160728.doc
  │   │   └── June
  │   │       ├── CA_CAPA.docx
  │   │       ├── CR_CAPA.docx
  │   │       ├── CZH_CAPA.docx
  │   │       ├── IN_CAPA.docx
  │   ├── Draft_Detail_Findings.xlsx
  │   └── trend.exe
  ├── trend.py
  └── write_data.py
  ```

* The Final CAPA's need to be saved in the newer format, .docx as Python-docx does not support .doc. The program will automatically
  convert .doc to .docx using wordconv.exe from the Microsoft Office Compatibility Pack and delete the older files.

  The Compatibility Pack can be downloaded here: https://www.microsoft.com/en-us/download/details.aspx?id=3

* Create the executable by forking this repo into a new folder and run this command within the folder

  `pyinstaller --clean -F trend.py`


 

## Observations
Certain values will have to be hardcoded in. This is because there are some inconsistencies in the Final CAPA Reports.

Inconsistent:
  - Two leads for one project, separated by "/", e.g. "Adam/Monika"
  - Batch name sometimes not included under "Project Information"
  - Customer name is sometimes the site name, other times it is site name + entity, e.g. "Wisconsin State Lottery"
  - Different formatting/spelling of go live date, e.g. "GO date or "Go live date" or "Go-Live Date"
  - "Detail of Findings" table is sometimes 2nd from last table or 3rd from last
    - some reports have extra color code table at the end
  - Different formatting of "Detail of Findings" table
  - Newer versions of CAPA reports put "Project Information" and "Project Stakeholders" into a table
  - Newer versions of CAPA reports have two extra non-empty lines on first page
    
Consistent:
  - Batch name on first page is always followed by Lead(s)'s name
  - followed by reported date
  - SAP ID, Go Live, and Site always included under "Project Information"
