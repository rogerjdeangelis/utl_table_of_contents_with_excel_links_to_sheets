Create a linked table of contents to several excel sheets

for output see two solutions
Manual Links:  https://www.dropbox.com/s/io35kia4aivmv32/toc.xlsx?dl=0
Automatic Links: https://www.dropbox.com/s/tk3oa2nj532k0if/cyn.xlsx?dl=0


  WORKING CODE ( Two solutions )

   1. MANUAL Hyperlinks (nned to open sheet and activate. One time only)

      data TOC;
          TABLE_OF_CONTENTS  = '=HYPERLINK("[d:\xls\toc.xlsx]Females!A1","Report Females")';
          output;
          TABLE_OF_CONTENTS  = '=HYPERLINK("[d:\xls\toc.xlsx]Males!A1","Report Males")';

      ods excel file="d:/xls/toc.xlsx"

      ods excel options(sheet_name="TOC");
      proc report data=TOC;
         title "TOC";
         cols TABLE_OF_CONTENTS;
         define TABLE_OF_CONTENTS /display;

      footnote '=HYPERLINK("[d:\xls\toc.xlsx]TOC!A1","Click here to return to the table of contents")';

      ods excel options(sheet_name="FEMALES");
      proc report data=sashelp.class(where=(sex="F"));
        title "FEMALES";

      ods excel options(sheet_name="MALES");
      proc report data=sashelp.class(where=(sex="M"));
         title "MALES";

   2. AUTOMATIC (from Cynthia at SAS - Table of contents is spread out. Cannot be collapsed automatically?)

     ods excel file="d:/xls/cyn.xlsx" style=statistical
     options(embedded_titles="yes" contents="yes" embedded_footnotes="yes");

     title "Table 1:  Review Status by Review Type FY 2017";
     footnote link="#'The Table of Contents'!a1"  "Return to TOC"; run;

     ods proclabel= "Table 1:  Review Status by Review Type FY 2017";
     proc tabulate data=sashelp.class contents=' ';
     class age sex;
     table sex=' ',age="Review Status"*(n*f=comma5. rowpctn*f=comma7.2)
            / box="Review Type" contents=' ';
     keylabel N='#' RowPctN='%';
     run;

     ods proclabel= "Table 2:  Review Cholesterol Status";
     proc tabulate data=sashelp.heart contents=' ';
     class chol_status sex;
     table sex=' ',chol_status="Review Status"*(n*f=comma5. rowpctn*f=comma7.2)
            / box="Review Type" contents=' ';
     keylabel N='#' RowPctN='%';
     run;

     ods proclabel= "Table 3:  Review BP Status";
     proc tabulate data=sashelp.heart contents=' ';
     class bp_status sex;
     table sex=' ',bp_status="Review Status"*(n*f=comma5. rowpctn*f=comma7.2)
            / box="Review Type" contents=' ';
     keylabel N='#' RowPctN='%';
     run;
     ods excel close;


HAVE
====

     WORK.CLASS total obs=19

     Obs    NAME       AGE    HEIGHT    WEIGHT

       1    Alfred      14     69.0      112.5
       2    Alice       13     56.5       84.0
       3    Barbara     13     65.3       98.0
      ...   ...         ...    ...       ...
      17    Ronald      15     67.0      133.0
      18    Thomas      11     57.5       85.0
      19    William     15     66.5      112.0

WANT
====

  My output d:/xls/toc.xlsx
  Cynthias is similar d:/xls/cyn.xlsx

  +------------------------------------+
  |                         A          |
  |------------------+-----------------|
1 |Detail Report of Females            |  if you mouse over the text
  |------------------------------------+   you will see the link
2 |Detail Report of Males              |
  +------------------+-----------------+

   [TOC]


SHEET [SEX=Males]

  ----------------------------------------------------+
  |     A      |     B      |    C       |    D       |
  ----------------------------------------------------+
1 |NAME        |    SEX     |  HEIGHT    |  WEIGHT    |
  +------------+------------+------------+------------+
2 | ALFRED     |     M      |    69      |  112.5     |
  +------------+------------+------------+------------+
   ...
  +------------+------------+------------+------------+
N | WILLIAM    |     M      |   66.5     |  112       |
  +------------+------------+------------+------------+

    Click here to return to the table of contents        Bac to TOC

[MALES]




SHEET [SEX=Females]

  ----------------------------------------------------+
  |     A      |     B      |    C       |    D       |
  ----------------------------------------------------+
1 | NAME       |    SEX     |  HEIGHT    |  WEIGHT    |
  +------------+------------+------------+------------+
2 | ALICE      |     F      |    69      |  112.5     |
  +------------+------------+------------+------------+
   ...
  +------------+------------+------------+------------+
N | BARBARA    |     F      |   66.5     |  112       |
  +------------+------------+------------+------------+

    Click here to return to the table of contents

[FEMALES]


Click here to return to the table of contents


*                _              _       _
 _ __ ___   __ _| | _____    __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \  / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/ | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|  \__,_|\__,_|\__\__,_|

;

data class;
  set sashelp.class(keep=name sex height weight);
run;quit;

*                                  _     _ _       _
 _ __ ___   __ _ _ __  _   _  __ _| |   | (_)_ __ | | _____
| '_ ` _ \ / _` | '_ \| | | |/ _` | |   | | | '_ \| |/ / __|
| | | | | | (_| | | | | |_| | (_| | |   | | | | | |   <\__ \
|_| |_| |_|\__,_|_| |_|\__,_|\__,_|_|   |_|_|_| |_|_|\_\___/

;

%utlfkil(d:/xls/toc.xlsx);
data TOC;
    TABLE_OF_CONTENTS  = '=HYPERLINK("[d:\xls\toc.xlsx]Females!A1","Report Females")';
    output;
    TABLE_OF_CONTENTS  = '=HYPERLINK("[d:\xls\toc.xlsx]Males!A1","Report Males")';
    output;
run;quit;

title;footnote;
* create an excel workbook with two sheets;
ods excel file="d:/xls/toc.xlsx" style=statistical
   options(embedded_titles="yes" embedded_footnotes="yes");

ods excel options(sheet_name="TOC");
proc report data=TOC;
   title "TOC";
   cols TABLE_OF_CONTENTS;
   define TABLE_OF_CONTENTS /display;
run;quit;

footnote '=HYPERLINK("[d:\xls\toc.xlsx]TOC!A1","Click here to return to the table of contents")';

ods excel options(sheet_name="FEMALES");
proc report data=class(where=(sex="F"));
  title "FEMALES";
run;quit;

ods excel options(sheet_name="MALES");
proc report data=class(where=(sex="M"));
   title "MALES";
run;quit;

ods excel close;
run;quit;

* ____            _   _     _
 / ___|   _ _ __ | |_| |__ (_) __ _
| |  | | | | '_ \| __| '_ \| |/ _` |
| |__| |_| | | | | |_| | | | | (_| |
 \____\__, |_| |_|\__|_| |_|_|\__,_|
      |___/
;

%utlfkil(d:/xls/cyn.xlsx);
ods excel file="d:/xls/tsttoc.xlsx" style=statistical
options(embedded_titles="yes" contents="yes" embedded_footnotes="yes");

title "Table 1:  Review Status by Review Type FY 2017";
footnote link="#'The Table of Contents'!a1"  "Return to TOC"; run;

ods proclabel= "Table 1:  Review Status by Review Type FY 2017";
proc tabulate data=sashelp.class contents=' ';
class age sex;
table sex=' ',age="Review Status"*(n*f=comma5. rowpctn*f=comma7.2)
       / box="Review Type" contents=' ';
keylabel N='#' RowPctN='%';
run;

ods proclabel= "Table 2:  Review Cholesterol Status";
proc tabulate data=sashelp.heart contents=' ';
class chol_status sex;
table sex=' ',chol_status="Review Status"*(n*f=comma5. rowpctn*f=comma7.2)
       / box="Review Type" contents=' ';
keylabel N='#' RowPctN='%';
run;

ods proclabel= "Table 3:  Review BP Status";
proc tabulate data=sashelp.heart contents=' ';
class bp_status sex;
table sex=' ',bp_status="Review Status"*(n*f=comma5. rowpctn*f=comma7.2)
       / box="Review Type" contents=' ';
keylabel N='#' RowPctN='%';
run;
ods excel close;



