# TrendingProg
The purpose of this program is to assist with doing monthly trending analysis. The Standards Compliance team at IGT compiles a
spreadsheet of data from projects each month. This allows managers to observe trends as they develop and help to better
identify problematic process areas and their causes. Getting the data for trending analysis is a very manual task that involves
extracting certain pieces of information in Final CAPA Reports and putting them into a spreadsheet. This program attempts to
automate as much of that process as possible.

# Observations
Certain values will have to be hardcoded in. This is because there are some inconsistencies in the Final CAPA Reports.

Inconsistent:
  - Two leads for one project, separated by "/", E.g. "Adam/Monika"
  - Batch name sometimes not included under "Project Information"
  - Customer name is sometimes the site name, other times it is site name + entity, E.g. "Wisconsin State Lottery"
  - Different formatting/spelling of go live date, E.g. "GO date or "Go live date" or "Go-Live Date"
  - "Detail of Findings" table is sometimes 2nd from last table or 3rd from last
    - some reports have extra color code table at the end
  - Different formatting of "Detail of Findings" table
    
Consistent:
  - Batch name is always 3rd non-empty line in document
  - followed by lead(s)'s name
  - followed by report date
  - SAP ID always included under "Project Information"
  
# 
