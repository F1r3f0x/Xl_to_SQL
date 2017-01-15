"""
        Patricio Labin Correa - 01/17

        ExlToSQL.py -
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

print('\n====================================================================='
        '==========\n')
print('exlToSQL - Transforma un archivo excel en entrada de datos para SQL' + 
        'standard\n')

# Si hay 2 argumentos (script + archivo) seguir
if len(sys.argv) == 2:  
    file_name = sys.argv[1] # asignar el 2do argumento como nombre del archivo
    start_time = time.clock()
    print ('Leyendo libro {} \n'.format(file_name))
    # Obtener el archivo excel y leer solo valores (no las formulas)
    book = pyxl.load_workbook(file_name,read_only = True, data_only = True)
    output_file = open("output.sql","w")
    output_file.write('/* Generado con exlToSQL.py -- Patricio Labin ' +
                            'Correa\n')
    output_file.write('Archivo = {}\n'.format(sys.argv[1]))
    output_file.write('===============================================' +
                            '=====================*/\n')

    for sheet in book.worksheets:  # Recorremos el libro por hoja (worksheets)
        print ('Leyendo Hoja = {}'.format(sheet.title))
        row_counter = 0
        for row in sheet.iter_rows():
            row_counter += 1
            if row_counter > 1:       # ignorar el titulo de la tabla
                sentence = ''.join(['insert into ',sheet.title,' values('])
                first_cell = True
                for celda in row:
                    #introducir el contenido de la celda en una string
                    cell_content = str(celda.value)
                    empty_cell = False   
                    if cell_content != 'None':      # ignorar celdas vacias
                        if not cell_content.isnumeric():
                            compare = cell_content.lower()
                            if not(compare == 'true' or compare == 'false'
                                    or compare == 'null'):
                                cell_content = ''.join(["'",cell_content,
                                                        "'"])
                        # introducir comas despues de la 1ra celda
                        if not first_cell:
                            sentence = ''.join([sentence,',',cell_content])
                        else:
                            first_cell = False
                            sentence = ''.join([sentence,cell_content])
                #escribir sentencia al archivo
                sentence = ''.join([sentence,');'])
                output_file.write('{}\n'.format(sentence))
         #filas procesadas sin el titulo
        output_file.write('\n')
        print('Filas procesadas = {}'.format(row_counter-1))
    output_file.close()
    print()
    print('Operacion Completada!, Tiempo de proceso = {}s'
            .format(time.clock()-start_time))
else:
    print ('Solo se permite un archivo')
print()
print ('====================================================================' +
        '===========')