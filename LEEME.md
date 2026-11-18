# ODS TEC — Inicio de Python: inventario verificable

Esta primera entrega revisa los archivos de extracción de Power Automate y crea una selección de correos para el análisis posterior, el programa lee los originales y escribe sus informes en una carpeta de salida independiente

El período predeterminado es **2026-06-10 a 2026-11-16**, correspondiente al histórico acordado, no cambia automáticamente al día de ejecución

## Preparar tu Mac

Instala Python 3.10 o posterior desde https://www.python.org/downloads/ si todavía no está instalado, después instala Visual Studio Code desde https://code.visualstudio.com/ y descomprime este paquete en una carpeta local de trabajo, por ejemplo una carpeta llamada ODS_TEC_Python

Abre esa carpeta desde VS Code mediante **Archivo → Abrir carpeta**, después abre su terminal integrada mediante **Terminal → Nueva terminal**

Comprueba Python:

```bash
python3 --version
```

Prepara el entorno del proyecto:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

El inventario utiliza únicamente la biblioteca estándar de Python, esta etapa no requiere instalar paquetes adicionales ni configurar una API de IA

Referencia oficial sobre entornos virtuales: https://packaging.python.org/en/latest/guides/installing-using-pip-and-virtual-environments/

## Preparar los archivos de entrada

La carpeta raíz que selecciones debe contener estas carpetas hermanas:

| Carpeta | Contenido esperado |
|---|---|
| 01_Correos | JSON y EML de las ejecuciones |
| 02_Adjuntos | Archivos originales descargados |
| 03_Control | Controles diarios control_YYYYMMDD.json |

También se admiten nombres como 01correos, 02adjuntos y 03control, se selecciona la carpeta **que contiene las tres**, no la carpeta de un correo ni el ZIP de descarga, si existe una carpeta adicional creada al descomprimir, navega hasta el nivel correcto

Comprueba que los archivos de OneDrive estén descargados y disponibles en el Mac, utiliza una copia local o una sincronización ya terminada, ejecuta el inventario cuando no haya otro proceso de extracción modificando ese conjunto

Los archivos de control son necesarios para confirmar los correos completados, sin ellos se genera un informe, pero los mensajes quedan sin verificación del control y no se seleccionan para análisis

## Ejecutar

Con el entorno activado y la terminal situada en la carpeta del programa, ejecuta:

```bash
python inventario_ods.py --raiz "/ruta/a/ODS_TEC_Extraccion_Historica"
```

Sustituye la ruta entre comillas por la ubicación real, una alternativa en la terminal del Mac es escribir `python inventario_ods.py --raiz `, dejar un espacio al final, arrastrar la carpeta de extracción desde Finder a la terminal y pulsar Enter, utiliza la carpeta que contiene 01, 02 y 03

El programa muestra su avance cada 250 adjuntos y termina imprimiendo el resumen y la ubicación de los informes, la lectura de huellas recorre los adjuntos por bloques de 1 MiB, no carga todos los gigabytes a la vez

La salida predeterminada es una carpeta nueva dentro de **resultados_inventario**, situada junto al proyecto cuando ejecutas desde su carpeta, cada ejecución utiliza una carpeta distinta para conservar los informes anteriores

Para cambiar las fechas:

```bash
python inventario_ods.py --raiz "/ruta/a/ODS_TEC_Extraccion_Historica" --inicio 2026-06-10 --fin 2026-11-16
```

Para elegir una carpeta de salida nueva:

```bash
python inventario_ods.py --raiz "/ruta/a/ODS_TEC_Extraccion_Historica" --salida "/ruta/local/informe_nuevo"
```

El programa rechaza una salida que ya contenga archivos o que esté dentro de las carpetas originales 01, 02 o 03

## Qué comprueba

Cada registro requiere un correo.id original, un estado exactamente Completo, una fechaConsulta válida dentro del período, un ID presente en el control diario del mismo buzón y carpeta, un EML no vacío y la localización de todos los adjuntos declarados como Guardado, además compara los tamaños de los adjuntos cuando el manifiesto aporta ese dato

Los registros Pendiente, EnProceso o sin estado Completo quedan fuera de la selección, de esta forma los intentos fallidos de los correos que decidiste omitir permanecen documentados sin considerarse completados, el resultado real del inventario determinará sus cantidades y no se fuerza un total de 933 ni de diez excluidos

La deduplicación de correos utiliza correo.id, la clave GUID del archivo sirve para encontrar sus archivos y no sustituye ese identificador, cuando varias copias completas del mismo ID contienen un correo idéntico se elige una sola, si el contenido de las copias completas difiere se solicita revisión y no se elige automáticamente una versión

Los adjuntos se identifican por las rutas guardadas en el manifiesto y se relacionan por su nombre de archivo local, nombres ambiguos en subcarpetas quedan pendientes, los archivos idénticos se agrupan por SHA-256 para revisar duplicación binaria, no se borra ni se desvincula ninguna copia

La comparación de internetMessageId entre distintos IDs genera una advertencia para revisar posibles copias del mensaje, no realiza una fusión automática, los archivos sin referencia localizada también se reportan y no se eliminan

La comprobación del EML en esta fase verifica presencia y tamaño no nulo, la integridad MIME completa y la concordancia de su contenido se revisarán al implementar la lectura, el inventario tampoco determina por sí solo qué PDF necesita OCR

## Informes generados

| Archivo | Uso |
|---|---|
| resumen_inventario.json | Conteos de archivos, IDs únicos, selección e incidencias |
| correos.csv | Un registro por archivo JSON y su estado de revisión |
| correos_seleccionados.csv | Una copia verificada por correo seleccionado |
| adjuntos.csv | Tamaños, extensiones, huellas y referencias |
| duplicados_correos.csv | Versiones que comparten el mismo ID original |
| duplicados_adjuntos.json | Archivos distintos con contenido binario idéntico |
| incidencias.csv | Registros no utilizables, archivos sin relación y avisos |
| seleccion_analisis.json | Rutas de los correos seleccionados para la próxima fase |

Los CSV usan UTF-8 con marca de codificación y separador punto y coma, si Excel los abre en una sola columna, impórtalos indicando ese separador, las incidencias son hallazgos por archivo y no representan necesariamente correos distintos

La selección reúne todos los criterios anteriores, un total menor que 933 requiere revisar motivos, controles y rutas antes de iniciar la IA, un archivo repetido no es por sí solo un error y la presencia de un mensaje en un control no demuestra que su copia local esté completa

## Exclusiones explícitas opcionales

Para excluir un correo que sí tenga un JSON Completo, crea un archivo local con esta estructura y añade los identificadores originales correspondientes:

```json
{
  "ids_correo": []
}
```

Después utiliza el parámetro `--exclusiones "/ruta/a/exclusiones.json"`, no añadas el GUID de un nombre de archivo, el programa no adivina los diez IDs omitidos a partir de un conteo

## Primer resultado que revisaremos

Comparte **resumen_inventario.json**, ese archivo permitirá contrastar los 1.906 elementos de 01_Correos, los 2.699 de 02_Adjuntos y los 933 correos reconocidos por el último resumen del flujo, si hay diferencias revisaremos después las filas pertinentes de los CSV

El paso siguiente será implementar la extracción de texto y OCR sobre una muestra de los correos seleccionados, después conectaremos la IA y validaremos la clasificación, esta entrega implementa el inventario y la selección local, todavía no ejecuta OCR, clasificación ni llamadas a una API

## Validación de esta entrega

El programa pasó 18 pruebas locales con datos sintéticos, cubriendo controles y parejas completas, reintentos, registros pendientes, IDs ausentes, adjuntos faltantes o ambiguos, tamaños distintos, versiones contradictorias, controles ausentes o incompatibles, exclusiones, archivos corruptos y preservación de los originales

Los 4,30 GB de tu extracción no se han recibido ni procesado aquí, la ejecución sobre tu conjunto real corresponde al Mac, para ejecutar las pruebas incluidas utiliza `python -m unittest -v` desde la carpeta del programa
