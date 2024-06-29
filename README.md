# Wind_Resume
The script is a summary generation tool in the context of the analysis of wind data from the Wind Explorer at https://eolico.minenergia.cl/

# Script

El script se describe de la siguiente forma.

 Lectura y preprocesamiento de datos:
 Lee datos de viento de un archivo CSV.
        Interpola valores faltantes.
        Filtra los datos por un rango de años específico (2007-2015).
        Convierte la columna de fecha/hora al formato adecuado.
        Agrega una columna para el mes.

    Cálculo de promedios:
        Calcula el promedio de la velocidad del viento por mes y hora.

    Exportación y formateo de resultados:
        Exporta los resultados a un archivo Excel.
        Reorganiza los datos en el Excel para tener un formato más legible.
        Redondea los valores a dos decimales.
        Agrega una columna con los nombres de los meses.

    Estilo condicional en Excel:
        Aplica bordes a todas las celdas.
        Colorea las celdas según rangos de velocidad del viento.
        Cuenta las celdas por color y mes, y muestra estos conteos en una tabla resumen.

# Instrucciones de uso (ESP).

Para poder generar un resumen (Mes/Hora) en formato .xlsx a partir de las series de viento disponibles en el explorador eólico, se deben seguir los siguientes pasos:

A) Ubicar geográficamente el punto en cuestión y acceder al menú ¨Explorar Recurso Eólico¨-
B) Seleccionamos la altura a la que se desea descargar la información del punto en cuestión.
D) Descargamos el archivo ¨Serie Horaria Viento Reconstruido 1980-2017.csv¨
E) El script y el conjunto de datos deben estar en la misma carpeta. Una vez confimado este punto, abrimos el script .py y cambiamos el parametro h (indica la altura de las mediciones elegida en el punto B).
F) Acomodamos los paths y ejecutamos el script.

El resultado debería ser un documento excel con el siguiente formato:
![imagen](https://github.com/naranguiz/Wind_Resume/assets/43880651/0d6974a3-c738-4835-b9c4-23238352e3f4)

# Significado de los valores y colores



