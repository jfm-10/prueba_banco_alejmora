Este script completo hace lo siguiente:

Configuración inicial: Importa las bibliotecas necesarias y configura el logging para su monitoreo y validaciones.
Conexión a la base de datos: Establece una conexión con la base de datos SQLite database.sqlite.
Carga de datos: Ejecuta una consulta SQL donde:

- Calcula las llamadas exitosas y no exitosas por empresa y mes.
- Aplica la lógica de cálculo de comisiones base para cada empresa.
- Calcula el porcentaje de descuento según las reglas específicas.

Procesamiento de datos:

- Calcula el descuento real.
- Ajusta el valor de comisión aplicando el descuento.
- Calcula el IVA y el total.
- Reordena las columnas para el formato final.


Exportar a Excel: Guarda los resultados en un archivo Excel.
Función principal: Orquesta todo el proceso, maneja errores y muestra un resumen de los resultados.

Para usar este script:

Tener instaladas las bibliotecas sqlite3, pandas, y openpyxl.
Ajustar la ruta de la base de datos (db_path) en la función main() si es necesario.
Ejecuta el script en tu entorno de Python.
