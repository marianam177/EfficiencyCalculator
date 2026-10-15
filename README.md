# Calculadora de eficiencia empresarial

## ¿Qué es esta carpeta?

Esta carpeta contiene la versión de escritorio de la calculadora para
Windows:

- `CalculadoraEficienciaDesktop.exe`: la aplicación.
- `Historial_Eficiencia.xlsx`: el historial activo de esta instalación.
- `README.md`: esta guía.

La aplicación tiene una ventana para registrar una evaluación, revisar
resultados y recomendaciones, consultar el historial y comparar dos
evaluaciones guardadas. Los gráficos de comparación se dibujan dentro de la
aplicación.

## Abrir la aplicación

1. Abre esta carpeta en el Explorador de archivos de Windows.
2. Haz doble clic en `CalculadoraEficienciaDesktop.exe`.
3. En la pestaña **Nueva evaluación**, completa los datos.
4. Pulsa **Calcular y guardar**.
5. Consulta los indicadores y las recomendaciones en **Resultados**.

No hace falta abrir Visual Studio Code, escribir comandos ni instalar Python
para usar el `.exe`. La aplicación fue preparada para Windows 10 y Windows 11.

## Ingresar una evaluación

Completa la empresa, el periodo y los datos de las seis secciones:

| Sección | Datos solicitados |
| --- | --- |
| Producción | Producción real y producción objetivo, en unidades comparables. |
| Calidad | Unidades producidas y unidades conformes. |
| Utilización de recursos | Horas disponibles y horas productivas. |
| Costos | Costo presupuestado y costo real, en la misma moneda y periodo. |
| Entregas | Total de entregas y entregas realizadas a tiempo. |
| Finanzas | Ventas totales y costos operativos del mismo periodo. |

El periodo se escribe como `AAAA-MM-DD`, por ejemplo `2026-10-14`. Al abrir
la aplicación, aparece la fecha actual como valor inicial; puedes cambiarla.

Escribe solo el valor numérico, sin signos de moneda ni separadores de miles.
Se aceptan decimales con punto o coma. Por ejemplo, puedes escribir `1250.5`
o `1250,5`.

La aplicación valida los campos obligatorios y comprueba, entre otras cosas,
que las unidades conformes no superen las producidas, que las entregas a
tiempo no superen el total y que las horas productivas no superen las
disponibles. Si un dato no es válido, muestra un aviso para corregirlo antes
de guardar.

## Resultados y fórmulas

Los indicadores se calculan con las fórmulas de la calculadora:

- **Productividad:** producción real / producción objetivo × 100.
- **Eficiencia de costos:** costo presupuestado / costo real × 100.
- **Calidad:** unidades conformes / unidades producidas × 100.
- **Cumplimiento de entregas:** entregas a tiempo / entregas totales × 100.
- **Utilización de recursos:** horas productivas / horas disponibles × 100.
- **Margen operativo:** (ventas totales − costos operativos) / ventas totales
  × 100.
- **Índice global:** promedio ponderado de los indicadores, según los pesos
  definidos en la calculadora. El margen operativo se lleva a una escala en
  la que un margen de 30 % equivale a 100 puntos para el índice.

La clasificación se determina así:

| Índice global | Clasificación |
| --- | --- |
| 90 % o más | EXCELENTE |
| 80 % a menos de 90 % | MUY EFICIENTE |
| 70 % a menos de 80 % | EFICIENTE / ACEPTABLE |
| 60 % a menos de 70 % | REQUIERE MEJORAS |
| Menos de 60 % | EFICIENCIA BAJA |

Las recomendaciones son alertas basadas en los umbrales programados. Son
orientativas: conviene interpretar los resultados de acuerdo con la actividad
y los criterios de cada empresa.

## Historial

El historial activo se llama siempre `Historial_Eficiencia.xlsx` y se guarda
en la misma carpeta que la aplicación. Por eso, en esta carpeta los cálculos
se guardan en el archivo de Excel que se encuentra aquí.

Cada evaluación exitosa se agrega como una fila nueva en la hoja
**Historial**. No se reemplazan las filas anteriores al guardar un cálculo
normal.

### Iniciar un historial nuevo sin perder el anterior

Si quieres comenzar un historial separado:

1. En **Nueva evaluación**, completa todos los datos.
2. Marca **Iniciar un historial nuevo y archivar el actual**.
3. Pulsa **Calcular y guardar** una sola vez.

La aplicación archiva el historial activo, crea el nuevo
`Historial_Eficiencia.xlsx` y guarda allí la evaluación que acabas de
calcular. El archivo archivado incluye la fecha en el nombre, por ejemplo
`Historial_Eficiencia_14_10_2026.xlsx`. Si ya existe un archivo con ese nombre,
se añade `_1`, `_2` y así sucesivamente, sin sobrescribir los anteriores.
La casilla se desmarca después de guardar correctamente.

La operación no se completa si no se puede guardar el nuevo archivo. Si
Excel tiene abierto el historial, ciérralo y vuelve a intentarlo.

### Consultar y comparar

- **Historial** muestra las evaluaciones del archivo activo. Usa **Actualizar**
  para volver a leerlo si cambió mientras la aplicación estaba abierta.
- **Comparar** permite seleccionar exactamente dos evaluaciones guardadas y
  muestra sus indicadores y un gráfico comparativo.
- El gráfico compara puntajes de 0 a 100. Para el margen operativo usa la
  misma escala de referencia del índice global; los valores negativos se
  limitan a 0 y los puntajes por encima de 100 se limitan a 100 en el índice.

Las evaluaciones que estén en archivos archivados no aparecen en las pestañas
**Historial** y **Comparar**. Esas pestañas leen el archivo activo; puedes
abrir un archivo archivado con Excel para consultarlo.

## Usar la aplicación en otro computador Windows

Para llevar la aplicación a otro equipo:

1. Copia la carpeta completa `app`, no solo el archivo `.exe`.
2. Pega la carpeta en una ubicación donde tengas permiso para guardar
   archivos, por ejemplo una carpeta personal o el Escritorio.
3. Abre el `.exe` desde esa carpeta.

La copia incluye el historial que existe en esta carpeta en el momento de
copiarla. A partir de entonces, cada copia mantiene su propio
`Historial_Eficiencia.xlsx`; los historiales no se sincronizan ni se comparten
automáticamente entre computadores. Si quieres pasar los cálculos a otro
equipo, transfiere también el archivo de historial de forma segura.

No hace falta una conexión a Internet para calcular, guardar o comparar
evaluaciones. Conserva el `.exe` y el Excel juntos en la carpeta para que la
aplicación utilice el historial que corresponde a esa instalación.

## Copias de seguridad

El archivo `Historial_Eficiencia.xlsx` contiene los registros de la
calculadora. Haz una copia de seguridad periódica, especialmente antes de
trasladar la carpeta, reinstalar Windows o iniciar un historial nuevo.
También puedes copiar los archivos `Historial_Eficiencia_DD_MM_AAAA.xlsx`
archivados a un dispositivo o ubicación segura.

No edites el historial mientras la aplicación o Excel lo tengan abierto. Si
Excel lo tiene bloqueado, cierra el archivo y vuelve a intentar guardar.

## Solución de problemas

### La aplicación parece no abrir

Espera unos segundos y revisa si la ventana se abrió detrás de otras
ventanas. Comprueba también que `CalculadoraEficienciaDesktop.exe` esté junto
al Excel del historial y vuelve a abrirlo.

Si Windows muestra una advertencia de seguridad, verifica que abriste el
ejecutable de esta carpeta y que sabes de dónde proviene antes de decidir si
quieres ejecutarlo.

### No se puede guardar el historial

Cierra `Historial_Eficiencia.xlsx` en Excel e inténtalo de nuevo. Comprueba
que la carpeta permita crear y modificar archivos y que haya espacio
disponible.

### No aparecen registros anteriores

La aplicación muestra solo las filas del historial activo ubicado en su
carpeta. Revisa que el archivo se llame exactamente
`Historial_Eficiencia.xlsx` y que contenga la hoja **Historial**. Los
historiales con fecha en el nombre son archivos archivados y no se combinan
automáticamente con el historial activo.

### Quiero empezar un historial nuevo

Marca la casilla de inicio de historial nuevo cuando vayas a guardar la
primera evaluación que debe quedar en el nuevo historial. La aplicación
archiva el historial que estaba activo y deja intactos los archivos
archivados previamente.

## Archivos que no hacen falta para ejecutar esta versión

Esta aplicación de escritorio funciona con el `.exe` y el archivo
`Historial_Eficiencia.xlsx`. No necesita el archivo Python, el manual, el
Excel de datos de ejemplo ni el CSV para ejecutarse.

## Información para mantenimiento

El `.exe` es una versión empaquetada para Windows de la aplicación gráfica.
Para ejecutar el código fuente o volver a compilarlo se requieren Python y
las dependencias de desarrollo correspondientes. Eso no es necesario para
el uso habitual de esta carpeta.
