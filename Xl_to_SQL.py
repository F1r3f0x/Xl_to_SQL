"""
        Patricio Labin Correa - 01/17

        Xl_to_SQL.py -
        Scans a formated excel file and creates an sql script to fill
        a database

        Dependencies:
        -   Python 3.6
        -   Openpyxl 2.4.1

        TODO:
        - Locate tables
        - Multiple file processing
        - Improve file input
"""

import sys
import time
import argparse
import openpyxl as pyxl

parser = argparse.ArgumentParser(prog='Xl_to_SQL.py')
parser.add_argument('file', type=str, help='Input Excel file to process')
args = parser.parse_args()

print('\n====================================================================='
        '==========\n')
print('Xl to SQL - Scans a formated excel file and creates an sql script to fill' +
        'a database\n')

file_name = args.file # 2nd arg is our file
start_time = time.clock()
print ('Scanning Book =  {} \n'.format(file_name))
# we open the excel file and read only values
book = pyxl.load_workbook(file_name,read_only = True, data_only = True)
output_file = open("output.sql","w")
output_file.write('/* Generated with Xl_to_SQL.py -- Patricio Labin ' +
                        'Correa\n')
output_file.write('File = {}\n'.format(sys.argv[1]))
output_file.write('===============================================' +
                        '=====================*/\n')

for sheet in book.worksheets:  # We start scanning the book sheet by sheet
    print ('Scanning Sheet = {}'.format(sheet.title))
    row_counter = 0
    for row in sheet.iter_rows():
        row_counter += 1
        if row_counter > 1:       # we don't care about the table title
            sentence = ''.join(['insert into ',sheet.title,' values('])
            first_cell = True
            for celda in row:
                # dump the cell content to a string
                cell_content = str(celda.value)
                empty_cell = False   
                if cell_content != 'None':      # we don't care about empty cells
                    if not cell_content.isnumeric():
                        compare = cell_content.lower()
                        if not(compare == 'true' or compare == 'false'
                                or compare == 'null'):
                            cell_content = ''.join(["'",cell_content,
                                                    "'"])
                    # here we put commas
                    if not first_cell:
                        sentence = ''.join([sentence,',',cell_content])
                    else:
                        first_cell = False
                        sentence = ''.join([sentence,cell_content])
            #write the final sql instruction to the output file
            sentence = ''.join([sentence,');'])
            output_file.write('{}\n'.format(sentence))
    output_file.write('\n')
    print('Rows proccesed = {}'.format(row_counter-1))
output_file.close()
print('\nOperation complete!, Processing Time = {}s'
        .format(time.clock()-start_time))
print ('\n====================================================================' +
        '===========')