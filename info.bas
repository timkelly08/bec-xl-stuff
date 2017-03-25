Attribute VB_Name = "info"
'''''''''''''''''''''''''''''''''''''''''''''
' Bechtel QTO Report Automation
' Assemble Systems
' By: Tim Kelly
' Version: 0.70
' Last Modified: 01/15/2017
'''''''''''''''''''''''''''''''''''''''''''''


'12/09/2016
'ADDED CHECK FOR PARENTHESIS IN COLUMN HEADERS DURING COLUMN VALIDATION SINCE ASSEMBLE UOM ON COLUMNS ARE INCONSISTENT BETWEEN PROJECTS
'ADDED ABILITY TO USE QUOTES IN ASSEMBLY CODE LVL 3 AND ITEM NAMES WITHIN THE QUANTITY FORMULA



'12/13/2016
'Version 0.4
'ADDED ABILITY TO CHECK FIRST FOR MATL ESTIMATE QUANTITY, THEN CHECK FOR SUBCONTRACTOR VALUE
'ADDED ERROR MSG AND RED CELL FILL FOR DUPLICATE QUANTITIES PULLED FROM THE COST REPORT
'FIXED CIRCULAR REFERENCE ISSUE BY CHANGING ITEM <> PITEM TO ITEMLONG <> PITEMLONG. ITEMLONG = ASSEMBLYCODE-ITEM


'01/11/2017
'Version 0.5
'ADDED COMPARISON ACROSS AREAS REPORT
'CHANGED VARIANCE FORMULA IN COMPARISON TAB

'01/11/2017
'Version 0.6
'CHANGED VARIANCE FORMULA IN COMPARISON TAB BASED ON CRITERIA FROM FRED

'01/13/2017
'Version 0.65
'CHANGED VARIANCE FORMULA IN COMPARISON TAB BASED ON MEETING WITH FRED
'CHANGED COLUMN HEADERS FOR COMPARISON TAB
'CHANGED OUT OF TOLERANCE TO WITHIN TOLERANCE AND REVERSED PM REPORT REQUIREMENTS
'CHANGED WITHIN TOLERANCE TO BE GREEN WHEN TRUE AND RED WHEN FALSE
'ADDED RED TO CELLS THAT DO NOT HAVE QUANTITES FOUND

'01/13/2017
'ADDED LINE ITEMS FOR COST CODES FOUND IN ESTIMATE THAT ARE NOT MODELED
'ADDED ABILITY TO DROP IN NEW ASSEMBLY CODE TABLE

'01/15/2017
'ADDED FLEXIBILITY TO FINDING COLUMNS IN COST REPORT

'01/20/2017
'Version 0.80
'ADDED RESPONSIBLE PARTY, QTO SOURCE, AND QTO METHOD COLUMNS TO THE REPORTS
'CORRECTED ASSEMBLY CODE COLUMN IN MASTERQTO
'ADDED ML TAKEOFFS TO THE SEARCH LIST FROM THE COST REPORT (MATL > ML > SC)

'01/23/2017
'Version 0.90
'CORRECTED COLUMNS IN COMPARISON ACROSS AREAS
'ADDED TREND GRAPHS

'01/25/2017
'Version 0.95
'UPDATE TO COMPARE ACROSS AREAS

'01/30/2017
'Version 0.96
'ADD ABILITY FOR CIVIL 3D MODELS TO SKIP THE COLUMN CHECK AND IMPORT COLUMNS INCLUDED, AND CREATE COLUMNS WHERE NEEDED
'INCLUDES SAMPLE OF NEW COMPARISON LAYOUT

'


