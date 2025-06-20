"""
Add Excel Formula using openpyxl

This script opens an Excel file (`number.xlsx`) and adds a formula to calculate the sum  
of cells B1 through B8 using `=SUM(B1:B8)`. It places the formula in cell B11  
and adds a label "Sum is" in cell A11.

Author : Sachin Kumar  
GitHub : https://github.com/sachin-py/excel-automation
"""

from openpyxl import load_workbook  # To load and modify Excel files

# Load the existing workbook and access the sheet
workbook = load_workbook('number.xlsx')
sheet = workbook['Sheet']

# Define the formula to sum values from B1 to B8
formula = '=SUM(B1:B8)'

# Write the label and formula into the sheet
sheet['A11'] = 'Sum is'      # Label in column A
sheet['B11'] = formula       # Formula in column B

# Save the updated workbook
workbook.save('number.xlsx')

# Close the workbook
workbook.close()
