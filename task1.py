from datetime import date, time
import xlsxwriter
import openpyxl
from openpyxl import Workbook

workbook = xlsxwriter.Workbook('data_validate.xlsx')
worksheet = workbook.add_worksheet()

header_format = workbook.add_format({'border': 1, 'bold': True, 'align': 'center'})

title = worksheet.write_row("B1:Q1", ["Assembly Parts", "Standard Parts", "Revision Number", "Status", "Part Number", "Weight", "Material", "Quantity", "Notes", "Date", "Designed By", "Detailed", "Approved By", "Custom scale", "Custom paper size", "Orientation"], header_format)

worksheet.set_column('B1:Q1', 15)


#For Column A

numbers = list(range(0,301))

for row_number, data in enumerate(numbers):
    
    worksheet.write(row_number, 0, data) 

worksheet.write('A1', "S.No")

#For Dropdown Columns

worksheet.data_validation('C2:C300', {'validate': 'list', 'source' : ['Yes', 'No']}) 
worksheet.data_validation('E2:E300', {'validate': 'list', 'source' : ['Released', 'Not Released']})
worksheet.data_validation('P2:P300', {'validate': 'list', 'source' : ['A0', 'A1', 'A2', 'A3', 'A4', 'A5']})
worksheet.data_validation('Q2:Q300', {'validate': 'list', 'source' : ['Portrait', 'Landscape']})

#For Date Column K

worksheet.data_validation('K2:K300', {'validate':'date', 'criteria': 'between', 'minimum':date(2000,1,1), 'maximum' : date(2023,1,1), 'input_title': 'Enter Date:', 'input_message': 'as YYYY/MM/DD'})

#For Integer input columns

worksheet.data_validation('D2:D300', {'validate': 'integer', 'criteria': 'between', 'minimum': 1, 'maximum': 1000, 'input_title': 'Enter an', 'input_message': 'INTEGER 1 to 1000', 'error_title': 'Input value is not valid!', 'error_message': 'It should be an integer between 1 and 1000'})
worksheet.data_validation('I2:I300', {'validate': 'integer', 'criteria': 'between', 'minimum': 1, 'maximum': 25, 'input_title': 'Enter an', 'input_message': 'INTEGER 1 to 25', 'error_title': 'Input value is not valid!', 'error_message': 'Quantity should be between 1 and 25'})

#For Float columns

worksheet.data_validation('G2:G300', {'validate': 'decimal', 'criteria': 'between', 'minimum': 1, 'maximum': 1000, 'input_title': 'Enter', 'input_message': 'FLOAT VALUE', 'error_title': 'Input value is not valid!', 'error_message': 'Must be a float number'})
worksheet.data_validation('O2:O300', {'validate': 'decimal', 'criteria': 'between', 'minimum': 1, 'maximum': 1000, 'input_title': 'Enter', 'input_message': 'As Value:Value', 'error_title': 'Input value is not valid!', 'error_message': 'Must be a ratio'})

workbook.close()    
