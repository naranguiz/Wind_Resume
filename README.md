# Wind_Resume
The script is a summary generation tool in the context of the analysis of wind data from the Wind Explorer at https://eolico.minenergia.cl/

# Description

El código comienza leyendo datos de viento de un archivo CSV, completando valores faltantes y filtrando por años específicos. Luego, convierte la fecha/hora y extrae el mes. A continuación, agrupa los datos por mes y hora para calcular promedios de velocidad del viento, por hora y para cada día de la semana. Estos resultados se exportan a un archivo Excel, donde se reorganizan, redondean según rangos de velocidad. Finalmente, el código aplica formato condicional, coloreando celdas según rangos de velocidad y generando un resumen de conteos de colores por mes.

# Instrucciones de uso (ESP).

Para poder generar un resumen en formato .xlsx a partir de las series de viento disponibles en el explorador eólico, se deben seguir los siguientes pasos:

A) Acceder al sitio web https://eolico.minenergia.cl/ , ubicar geográficamente el punto en cuestión y acceder al menú ¨Explorar Recurso Eólico¨
B) Seleccionamos la altura deseada en la barra movible, esta será la altura de los datos que se descargarán en la ¨Serie Horaria Viento Reconstruido 1980-2017.csv¨.
D) Descargamos el archivo ¨Serie Horaria Viento Reconstruido 1980-2017.csv¨ para la ubicación que seleccionamos.
E) Abrimos el script .py, modificamos los paths y cambiamos el parametro h (este parametro indica la altura de las mediciones elegida en el punto B).


El resultado debería ser un documento excel con el siguiente formato:
![imagen](https://github.com/naranguiz/Wind_Resume/assets/43880651/0d6974a3-c738-4835-b9c4-23238352e3f4)

# Significado de los valores y colores



