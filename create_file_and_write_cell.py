"""
Excel Cell Writer

This script creates a new Excel file (output.xlsx), writes the value 'Hello' 
to cell B2 (row 2, column 2), and saves the workbook using openpyxl.

Author : Sachin Kumar  
GitHub : https://github.com/sachin-py/pattern_programs/
"""

from openpyxl import Workbook  # Import to create a new Excel workbook

# Create a new workbook
workbook = Workbook()

# Get the default active sheet
sheet = workbook.active

# Write 'Hello' to cell B2 (row=2, column=2)
sheet.cell(row=2, column=2, value='Hello')

# Save the workbook to a file named 'output.xlsx'
workbook.save('output.xlsx')

# Close the workbook
workbook.close()
