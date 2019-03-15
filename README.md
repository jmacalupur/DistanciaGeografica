# DistanciaGeográfica
Función que puedes agregar a tu Excel para calcular la distancia entre dos coordenadas geográficas por medio de la Fórmula del Haversine.

Comparto un pequeño módulo que podrás insertar a tu Archivo Excel cuando requieras calcular la distancia entre dos coordenadas geográficas. 

##¿Qué es la fórmula del Haversine?

Todo lo puedes leer desde [Wikipedia](https://es.wikipedia.org/wiki/Fórmula_del_haversine).

De manera resumida, se puede estimar la distancia mediante la siguiente ecuación:

![alt text](https://github.com/jmacalupur/DistanciaGeografica/blob/develop/imagenes/haversineFormula.png)


##Consideraciones

- Las Latitudes y Longitudes deben de estar en grados decimales. (No en grados minutos o segundos)
- Considerar el radio de la tierra en km, y según tu ubicación:
    - Para los polos: 6,356.8 
    - Para la zona ecuatorial: 6,378.10
    - Radio Promedio 6,371.0
- El orden de los datos en la función es la siguiente: (Latitud1, Longitud1, Latitud2, Longitud2, RadioTierra). Donde:
    - latitud1 : Latitud de tu punto inicial.
    - Longitud1: Longitud de tu punto inicial.
    - Latitud2 : Latitud de tu punto final.
    - Longitud2: Longitud de tu punto final.
    - RadioTierra: Radio de la tierra en Km.
- El resultado obtenido lo tendremos en metros.

##¿Cómo usar la función?

1. Haremos el procedimiento con un ejemplo.
2. Primero abrimos el archivo Excel donde veremos los datos de la latitud, longitud de los puntos inicial y final, y del radio de la tierra: 

![alt text](https://github.com/jmacalupur/DistanciaGeografica/blob/develop/imagenes/jmacalupurDistanciaGeografica_1.PNG)

3. Vamos a la pestaña Programador (tienes que tenerla activada), y vamos a Visual Basic:

![alt text](https://github.com/jmacalupur/DistanciaGeografica/blob/develop/imagenes/jmacalupurDistanciaGeografica_2.PNG)

4. Seleccionamos nuestro archivo dentro de la ventana y luego vamos a Archivo->Importar:

![alt text](https://github.com/jmacalupur/DistanciaGeografica/blob/develop/imagenes/jmacalupurDistanciaGeografica_3.PNG)

![alt text](https://github.com/jmacalupur/DistanciaGeografica/blob/develop/imagenes/jmacalupurDistanciaGeografica_4.PNG)

5. Luego buscamos el archivo "FuncionDistanciaGeográfica.bas" que descargamos del repositorio:

![alt text](https://github.com/jmacalupur/DistanciaGeografica/blob/develop/imagenes/jmacalupurDistanciaGeografica_5.PNG)

6. Si vamos a la carpeta Módulos, podremos ver el archivo cargado:

![alt text](https://github.com/jmacalupur/DistanciaGeografica/blob/develop/imagenes/jmacalupurDistanciaGeografica_6.PNG)


7. Luego vamos al Excel, y ejecutamos la función DistanciaGeografica. Recuerda el orden que te indico en la sección de **Consideraciones**:

![alt text](https://github.com/jmacalupur/DistanciaGeografica/blob/develop/imagenes/jmacalupurDistanciaGeografica_7.PNG)

8. Al final, damos Enter y obtendremos el resultado:

![alt text](https://github.com/jmacalupur/DistanciaGeografica/blob/develop/imagenes/jmacalupurDistanciaGeografica_8.PNG)


Eso es todo. 

