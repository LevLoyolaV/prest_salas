Completar formulario PRESTAMO DE SALAS

Descripción general y funcionalidad:

Este script es un programa que usa Python para leer y escribir información en un archivo de Excel, específicamente el documento de utilizado por CAD "PRESTAMO DE SALAS".
El programa carga un archivo de Excel específico en el que se desea trabajar.
A continuación, busca en la hoja de trabajo correspondiente y obtiene los datos que ya se encuentran en la tabla, para poder visualizar lo que ya está ingresado en el formulario.
Luego, el programa le pide al usuario que ingrese información adicional, a través de inputs, para agregarla a la tabla en el archivo de Excel. Esta información incluye la fecha, el nombre completo, la carrera, el RUT del responsable de la sala, la hora de inicio y término, la sala asignada, la actividad y el número de asistentes.
A continuación, el programa busca la primera fila vacía en la tabla, para poder agregar la información proporcionada por el usuario. Una vez encontrada la fila vacía,
el programa escribe los datos proporcionados por el usuario en las columnas correspondientes de esa fila.
Por último, el programa guarda los cambios realizados en el archivo de Excel original.


Descripción detalla del script

01 - Importa los módulos necesarios: openpyxl, que permite trabajar con archivos de Excel, y PrettyTable, que permite crear tablas en la terminal.
02 - Establece la ruta del archivo de Excel que se utilizará para almacenar los datos de préstamo de salas.
03 - Abre el archivo de Excel y selecciona la hoja de trabajo que contiene la tabla de registro de préstamo de salas.
04 - Crea una lista vacía llamada tabla_datos.
05 - Itera a través de las filas de la tabla de registro de préstamo de salas (filas 7 a 258) y agrega cada fila a la lista tabla_datos.
06 - Crea una tabla PrettyTable y agrega los nombres de las columnas.
07 - Agrega cada fila de la lista tabla_datos a la tabla PrettyTable.
08 - Imprime la tabla en la terminal.
09 - Pide al usuario que ingrese los detalles del préstamo de la sala, como la fecha, el nombre completo, la carrera, etc.
10 - Busca la primera fila vacía en la columna A de la tabla de registro de préstamo de salas.
11 - Si no se encuentra una fila vacía en la columna A, usa la siguiente fila después de la última fila llena.
12 - Escribe los valores ingresados por el usuario en las celdas correspondientes en la primera fila vacía encontrada.
13 - Alinea el contenido de cada celda para que esté centrado horizontal y verticalmente.
14 - Guarda los cambios en el archivo de Excel.
15 - Finalmente, imprime un mensaje en la terminal para indicar que se han guardado los datos en el archivo de Excel.
