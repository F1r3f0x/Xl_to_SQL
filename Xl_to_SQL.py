"""
    Patricio Labin Correa (F1r3f0x) - 02/18

    Xl_to_SQL.py -
    Scans a formatted excel file and creates an sql script to fill
    a database

    Dependencies:
    -   Python 3.6
    -   Openpyxl 2.4.5

    TODO:
    - Locate tables in a single page.
    - Multiple file processing
    - Parse SQL functions
"""

import time
import argparse
import openpyxl as pyxl


def get_sql_insert_sentence(table_name: str, values: list):
    """
    Creates an SQL sentence.
    :param table_name: table name for insert
    :type table_name: str
    :param values: values to insert
    :type values: list
    :return: SQL sentence
    :rtype: str
    """

    sql = f'INSERT INTO {table_name} VALUES ('

    parsed_values = []

    # Validate every value
    for k, val in enumerate(values):

        if val == 'None':
            parsed_values.append('\" \"')

        else:
            try:
                # It works ...
                float(val)
                val = str(val)
            except ValueError:
                comp_val = val.strip().lower()

                if not (comp_val == 'true' or comp_val == 'false'
                        or comp_val == 'null'):
                    val = f'\"{val}\"'
                else:
                    val = val.upper()

            parsed_values.append(val)

        # Commas only after the first value
        if k < len(values) - 1:
            parsed_values.append(',')

    # Merge all
    return ''.join((sql, *parsed_values, ');'))


if __name__ == '__main__':

    parser = argparse.ArgumentParser(prog='Xl_to_SQL.py')
    parser.add_argument('input_file', type=str,
                        help='Input Excel file to process')
    parser.add_argument('-o', type=str, default='output.sql',
                        help='Output SQL script')
    args = parser.parse_args()

    start_time = time.clock()

    print('\n================================================================' +
          '===============\n')
    print('Xl to SQL - Scans a formatted excel file and creates an sql script' +
          ' to fill a database\n')

    # we open the excel file and read only values
    file_name = args.input_file
    print('Scanning Book =  {} \n'.format(file_name))
    book = pyxl.load_workbook(file_name, read_only=True, data_only=True)

    output_file = open(args.o, "w", encoding="utf-8")
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
                row_values = []
                for i, cell in enumerate(row):
                    cell_content = cell.value
                    str_cell = str(cell.value)

                    row_values.append(str_cell)

                # write the final sql instruction to the output file
                sentence = get_sql_insert_sentence(sheet.title, row_values)
                output_file.write('{}\n'.format(sentence))

        output_file.write('\n')
        print('Rows proccesed = {}\n'.format(row_counter - 1))

    output_file.close()
    print('\nOperation completed!, Time = {}s'.format(time.clock() - start_time))
    print('\n===============================================================' +
          '================')
