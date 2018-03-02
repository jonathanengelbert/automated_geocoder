#==============================================================================

#NUMERATE

# Last Modified: 01/16/2018
# Author: Jonathan Engelbert (Jonathan.Engelbert@sfgov.org)

# Description: This script stores a dictionary that associates Excel column
# names with numbers (as in C = 3). It also stores the function called in
# imperfect_addresses.py that finds the target column

#==============================================================================

from string import ascii_lowercase
import openpyxl


excel_columns = {}
a_columns = {}
b_columns = {}
c_columns = {}


#Numerates alphabet, 4 times:
i = 1
for c in ascii_lowercase:
    excel_columns[c] = i
    a_columns['a' + c] = 26 + i
    b_columns['b' + c] = 52 + i
    c_columns['c' + c] = 78 + i
    i += 1

excel_columns.update(a_columns)
excel_columns.update(b_columns)
excel_columns.update(c_columns)


#FUNCTIONS:

def get_address_column(workbook):
    """Iterates through wb sheet to find the column with addresses to be 
    transformed. Looks for either 'Address' or 'Location' headers"""
    wb = workbook
    ws = wb.active
    # Assigns the values retrived from cells to variable "value:
    value = (ws.cell)
    for col in ws.iter_cols(min_row=1, max_col=100, max_row=1):
        for cell in col:
            if(cell.value):
                if "Address" in (cell.value) or "address" in (cell.value):
                    target = cell.column.lower()
                    for k, v in excel_columns.items():
                        if k == target:
                            target = v
                            return target


