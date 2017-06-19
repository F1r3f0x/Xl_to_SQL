"""
    Patricio Labin Correa (F1r3f0x) - 01/17

    Xl_to_SQL.py -
    Scans a formated excel file and creates an sql script to fill
    a database

    Dependencies:
    -   Python 3.5
    -   Openpyxl 2.4.1

    TODO:
    - Locate tables
    - Multiple file processing
    - Improve file input
    - Transform excel date to sql date (ready)
    - Handle datetime
"""

import time
import argparse
import openpyxl as pyxl

if __name__ == '__main__':

    parser = argparse.ArgumentParser(prog='Xl_to_SQL.py')
    parser.add_argument('Input_File', type=str,
                        help='Input Excel file to process')
    parser.add_argument('--output', type=str, default='output.sql',
                        help='Output SQL script')
    args = parser.parse_args()

    start_time = time.clock()

    print('\n================================================================' +
          '===============\n')
    print('Xl to SQL - Scans a formated excel file and creates an sql script' +
          ' to fill a database\n')

    # we open the excel file and read only values
    file_name = args.Input_File
    print('Scanning Book =  {} \n'.format(file_name))
    book = pyxl.load_workbook(file_name, read_only=True, data_only=True)

    output_file = open(args.output, "w", encoding="utf-8")
    output_file.write('/* Generated with Xl_to_SQL.py -- Patricio Labin ' +
                      'Correa\n')
    output_file.write('===============================================' +
                      '=====================*/\n')

    for sheet in book.worksheets:  # We start scanning the book sheet by sheet
        print('Scanning Sheet = {}'.format(sheet.title))
        row_counter = 0
        for row in sheet.iter_rows():
            row_counter += 1
            if row_counter > 1:       # we don't care about the table title
                sentence = ''.join(['insert into ', sheet.title, ' values('])
                first_cell = True
                for celda in row:
                    # dump the cell content to a string
                    cell_content = celda.value
                    str_cell = str(celda.value)
                    empty_cell = False
                    # we don't care about empty cells
                    if cell_content != 'None':
                        """
                        #SQL date format
                        if isinstance(cell_content,datetime):
                            str_cell = cell_content.year + '-'
                            str_cell += cell_content.month + '-'
                            str_cell += cell_content.day
                        """
                        if not str_cell.isnumeric():
                            compare = str_cell.lower()
                            if not(compare == 'true' or compare == 'false'
                                   or compare == 'null'):
                                str_cell = ''.join(['"', str_cell, '"'])

                        # here we put commas
                        if not first_cell:
                            sentence = ''.join([sentence, ',', str_cell])
                        else:
                            first_cell = False
                            sentence = ''.join([sentence, str_cell])

                # write the final sql instruction to the output file
                sentence = ''.join([sentence, ');'])
                output_file.write('{}\n'.format(sentence))

        output_file.write('\n')
        print('Rows proccesed = {}\n'.format(row_counter - 1))

    output_file.close()
    print('\nOperation complete!, Processing Time = {}s'.format(time.clock() -
                                                                start_time))
    print ('\n===============================================================' +
           '================')