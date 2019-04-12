# Macros-Excel
Macro de similaridad de textos. Según tabla de porcentaje de similaridad.

Hola estimados remotos.

11/04/2019

EJEMPLO:
-------

COMPARACIÓN ENTRE DOS PALABRAS, a la izquierda el "modelo", a la derecha el "comparado"

FUNCIÓN:
fSimilar2 ("Rotonda","Redondas")

RESULTADO

56        (% DE SIMILARIDAD)


EXPOSICIÓN:
----------

A los que ya saben de macros (VBA) pueden abrir el documento adjunto, y probar en el panel de Inmediato.
?fSimilar2("Rotonda","Redondas")
El resultado en el mismo panel debe ser el siguiente.

Rotonda       Redondas      Escala: 114%
1(14%), 1(14%), 4(57%),   - mayor : 8
Puntaje 1: 56
Redondas      Rotonda       Escala: 88%
1(14%), 1(14%), 1(14%), 3(43%), 
Puntaje 2: 31
 56 

La función fSimilar2 imprime resultados de cada sección medible de sus comandos.
Para poder utilizar el procedimiento sin presentar estos datos debe utilizar la función del ejemplo siguiente.
fSimilar("Rotonda","Redondas")
El resultado devuelto por la función debe ser el siguiente.
56

PASO A PASO: (Ejecución de la prueba)
-----------
Para quienes no saben ejecutar comandos en macros de excel (VBA).
Abrir el documento adjunto Equivalecia de Campos.xlsm (La XLSM indica que contiene código VBA)
Con el documento abierto presione Alt + F11
En la ventana que aparece que es el editor de Visual Basic Para Aplicaciones.
Presione en el Menú Ver, submenú Ventana Inmediato. O presione Ctrl + G.
En el panel que aparece en la parte inferior con el título Inmediato.
Escriba el siguiente código: "?fSimilar2("Rotonda","Redondas")"
Sin comillas e incluyendo el signo de cerrar interrogación.
El resultado debe verse como en la explicación antes de esta sección.

EXPLICACIÓN:
-----------

En el la comparación de ambos textos se convierten los caracteres a mayúsculas para simplificar la comparación.
ROTONDA
REDONDAS

Se mide la escala que diferencia cada muestra
ROTONDA>REDONDAS         Escala: 114%
REDONDAS>ROTONDA         Escala: 88%

Quiere decir que el tamaño de REDONDAS es mayor que ROTONDA con 114%
Y el tamaño de ROTONDA es el 88% de REDONDAS



... continuará
