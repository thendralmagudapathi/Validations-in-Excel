from datetime import date
import xlsxwriter

workbook = xlsxwriter.Workbook('test1.xlsx')
worksheet = workbook.add_worksheet()

header_format = workbook.add_format({'border': 1})

row = 0
col = 0

heading_list = ["S.no","Assembly Parts","Standard Parts","Revision Number","Status","Part Number","Weight","Material","Quantity","Notes","Date","Designed By","Detailed","Approved By","Custom Scale","Custom Paper Size",
"Orientation"]
for j, t in enumerate(heading_list):
    worksheet.write(row, col + j, t, header_format)

worksheet.data_validation("C2:C300",
                    {'validate': 'list',
                    'source': ['Yes','No']})

worksheet.data_validation("D2:D300",
                            {'validate': 'integer',
                            'criteria': '>',
                            'value': 0,
                            'input_title': 'Enter an integer',
                            'input_message': 'between 1 and 100',
                            'error_title': 'Input value is not valid!',
                            'error_message':'It should be an integer between 1 and 100'},
                            )

worksheet.data_validation("E2:E300",
                    {'validate': 'list',
                    'source': ["Released","Not Released"]})

worksheet.data_validation('G2:G300', {'validate': 'decimal',
                                'criteria': '>',
                                'value': 0.1,
                                'input_title': 'Enter an Float:',
                                'error_title': 'Input value is not valid!',
                                'error_message':'It should be an integer between 0.1 and 0.10',
                                })


worksheet.data_validation("I2:I300",
                            {'validate': 'integer',
                            'criteria': '>',
                            'value': 0,
                            'input_title': 'Enter an integer:',
                            'error_title': 'Input value is not valid!',
                            'error_message':'It should be an integer between 1 and 100',
                            }

                           )


worksheet.data_validation('K2:K300', {'validate': 'date',
                                 'criteria': 'between',
                                 'minimum': date(2013, 1, 1),
                                 'maximum': date(2024, 12, 12),
                                 
                                 })

worksheet.data_validation('O2:O300',{'validate': 'decimal',
                                'criteria': '>',
                                'value': 0.1,
                                'input_title': 'Enter an Float:',
                                'input_message':'Enter the values between 0.1 to 0.10',
                                'error_title': 'Input value is not valid!',
                                'error_message':'It should be an integer between 0.1 and 0.10',
                                })

worksheet.data_validation("P2:P300",
                    {'validate': 'list',
                    'source': ["A0","A1","A2","A3","A4"]})

worksheet.data_validation("Q2:Q300",
                    {'validate': 'list',
                    'source': ["Landscape","Portrait"]})

workbook.close()