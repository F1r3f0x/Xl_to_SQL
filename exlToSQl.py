# Patricio Labin Correa
"""
        TODO:
        - Reconocer Titulo de tabla
        - Procesar multiples archivos
        - Procesar nombres erroneos
        - Mejorar proceso de ingreso del archivo
"""

import sys
import time
import openpyxl as pyxl

print()
print('====================================================================='
        '==========\n')
print('exlToSQL - Transforma un archivo excel en entrada de datos para SQL' + 
        'standard\n')

# Si hay 2 argumentos (script + archivo) seguir
if len(sys.argv) == 2:  
    nombreArchivo = sys.argv[1] # asignar el 2do argumento como nombre del archivo
    tiempo_0 = time.clock()
    print ('Leyendo libro {} \n'.format(nombreArchivo))
    # Obtener el archivo excel y leer solo valores (no las formulas)
    libro = pyxl.load_workbook(nombreArchivo,read_only = True, data_only = True)
    archivoSentencias = open("output.sql","w")
    archivoSentencias.write('/* Generado con exlToSQL.py -- Patricio Labin ' +
                            'Correa\n')
    archivoSentencias.write('Archivo = {}\n'.format(sys.argv[1]))
    archivoSentencias.write('===============================================' +
                            '=====================*/\n')

    for hoja in libro.worksheets:  # Recorremos el libro por hoja (worksheets)
        print ('Leyendo Hoja = {}'.format(hoja.title))
        contadorFilas = 0
        for fila in hoja.iter_rows():
            contadorFilas += 1
            if contadorFilas > 1:       # ignorar el titulo de la tabla
                sentencia = ''.join(['insert into ',hoja.title,' values('])
                primeraCelda = True
                for celda in fila:
                    #introducir el contenido de la celda en una string
                    contenidoCelda = str(celda.value)
                    celdaVacia = False   
                    if contenidoCelda != 'None':      # ignorar celdas vacias
                        if not contenidoCelda.isnumeric():
                            comparar = contenidoCelda.lower()
                            if not(comparar == 'true' or comparar == 'false'
                                    or comparar == 'null'):
                                contenidoCelda = ''.join(["'",contenidoCelda,
                                                        "'"])
                        # introducir comas despues de la 1ra celda
                        if not primeraCelda:
                            sentencia = ''.join([sentencia,',',contenidoCelda])
                        else:
                            primeraCelda = False
                            sentencia = ''.join([sentencia,contenidoCelda])
                #escribir sentencia al archivo
                sentencia = ''.join([sentencia,');'])
                archivoSentencias.write('{}\n'.format(sentencia))
         #filas procesadas sin el titulo
        archivoSentencias.write('\n')
        print('Filas procesadas = {}'.format(contadorFilas-1))
    archivoSentencias.close()
    print()
    print('Operacion Completada!, Tiempo de proceso = {}s'
            .format(time.clock()-tiempo_0))
else:
    print ('Solo se permite un archivo')
print()
print ('====================================================================' +
        '===========')