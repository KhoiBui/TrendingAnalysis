# TrendingProg

The purpose of this program is to assist with doing monthly trending analysis. The Standards Compliance team
maintains statistical data on every project and analyze them on a monthly, quarterly, and yearly basis. This
data is observed for trends that can help with identifying problematic areas in the software development life
cycle. A lot of the work behind doing this analysis is currently manual.

Some observations:

Certain values in the program will have to be hard-coded in. This is because there are some inconsistencies
in the Final CAPA reports.

Inconsistent:
* Two leads for one project, separated by "/" E.g. "Adam/Monika"
* Batch name sometimes not included under "Project Information"
* Customer name is sometimes just the site name, other times it is site name + entity
* Different formatting/spelling of go-live date E.g. "Go date" or "Go live date" or "Go-Live date"
* "Detail of Findings" table is sometimes 2nd from last table or 3rd from last
  - some reports have extra color code table at the end
  
Consistent:
* Batch name is always 3rd non-empty line in document
* followed by lead's name
* followed by report date
* SAP ID
* "Project Stakeholders" section always arranged in order
* "Detail of Findings" table have consistent header
  - "Proces Area"    "Goal"    "Practice"    "Description"    "Rating"
