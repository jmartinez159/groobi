# Groobi
This program is for automating the process of finding significant changes in rows in excel spreadsheets. Program clears filters first then reads in data from the last 3 sheets, ignoring the very last one(PM Sheet). Columns CUSTOMER CODE, SAP CODE, DESCRIPTION, PO, SAP ORDER, PRODUCED QTY, INVOICE, ESTIMATED DELIVERY DATE are used to track changes. 

# BEFORE RUNNING
-FILE NAME NEEDS TO BE "test1"
-FILE NEEDS TO BE IN dist folder

# Dependencies
PyInstaller, Pandas, Openpyxl, and Regular Expressions

# Possible Bugs
1. The name of the sheets in excel might overlap if they keep getting added to this one sheet.
2. Duplicates exist in data set and we are choosing to ignore them after the first appearance.

# Requirements
When differences are detected then reply with customer order and sap order
