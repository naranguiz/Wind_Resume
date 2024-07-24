# Análisis de viento

El script es una herramienta de generación de resúmenes en el contexto del análisis de datos de viento del Wind Explorer en https://eolico.minenergia.cl/

# Description

El código comienza leyendo datos eólicos presentes en un archivo CSV, completando valores faltantes y filtrando por años específicos. Luego, convierte la fecha/hora y extrae el mes. A continuación, agrupa los datos por mes y hora para calcular promedios de velocidad del viento, por hora y para cada día de la semana. Estos resultados se exportan a un archivo Excel, donde se reorganizan y redondean según rangos de velocidad descritos por tabla. Finalmente, el código aplica formato condicional, coloreando celdas según magnitudes de velocidad y generando un contador de las diferentes magnitudes presentes por mes.

# Significado de los valores y colores
Los colores del formato condicional representan magnitudes de viento redondeadas según rangos descritos por la siguiente tabla:

![imagen](https://github.com/user-attachments/assets/01c38439-ac81-43ee-bfe5-b5f498613e0b)

# Instrucciones de uso (ESP).

Para poder generar un resumen en formato .xlsx a partir de las series de viento disponibles en el explorador eólico, se deben seguir los siguientes pasos:

A. Acceder al sitio web https://eolico.minenergia.cl/ , ubicar geográficamente el punto en cuestión y acceder al menú ¨Explorar Recurso Eólico¨.
B. Seleccionamos la altura deseada en la barra movible, esta será la altura de los datos que se descargarán en la ¨Serie Horaria Viento Reconstruido 1980-2017.csv¨. \n
C. Descargamos el archivo ¨Serie Horaria Viento Reconstruido 1980-2017.csv¨ para la ubicación que seleccionamos. \n
D. Abrimos el script .py, modificamos los paths y cambiamos el parametro h (este parametro indica la altura de las mediciones elegida en el punto B). \n

El resultado debería ser un documento excel con el siguiente formato:
![imagen](https://github.com/naranguiz/Wind_Resume/assets/43880651/0d6974a3-c738-4835-b9c4-23238352e3f4)










