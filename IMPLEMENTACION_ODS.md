# Procesamiento ODS TEC

Power Automate guarda el correo, sus adjuntos y el control diario. Python
reúne la evidencia por correo y exporta a Excel. La clasificación con OpenAI
se habilita explícitamente y requiere autorización institucional para enviar
el texto de correos y adjuntos. La extracción por sí sola es local.

## Prueba local sobre el histórico

Desde la carpeta del repositorio y con la terminal del Mac:

```bash
source .venv/bin/activate
python -m pip install -r requirements-ods.txt
python analizar_ods.py \
  --seleccion resultados_inventario/20261118T205716004269Z/seleccion_analisis_verificada.json \
  --salida resultados_ods/historico --limite 25
```

El manifiesto debe tener al lado el `adjuntos.csv` de su inventario.
Conserva los datos en la ubicación registrada por el manifiesto. Para el
Mac se usan los comandos de Poppler, Tesseract con español e inglés, y
`sips` para convertir el primer fotograma GIF. Si faltan Poppler o Tesseract,
usa `brew install poppler tesseract-lang`.

El Excel se guarda como `resultados_ods/historico/Actividades_ODS.xlsx`.
Las hojas Control y Adjuntos muestran el avance de extracción. En esta
ejecución Actividades se mantiene vacía hasta habilitar la clasificación.
El texto de la evidencia queda en `progreso.sqlite3`.

## Clasificar la muestra

Guarda en `config_local/` los tres archivos del TEC. El script lee las
169 metas de `Metas ODS`, los 17 tipos de `TiposActividad` y consulta
localmente `Funcionarios_Directorio_TEC(1).xlsx`. Los libros originales
sirven como insumos. El resultado se escribe en un Excel independiente.

La asignación exige citas del correo o sus adjuntos. El directorio confirma
coincidencias únicas por nombre o correo y completa los datos de contacto.
La dependencia de adscripción en el directorio tiene su propia columna.
El remitente no se asigna automáticamente como responsable. Las coincidencias
ambiguas y las actividades con menos de dos metas comprobadas quedan para
revisión. Se exportan hasta diez códigos distintos por actividad consolidada.

Configura `OPENAI_API_KEY` en la sesión de terminal mediante una entrada
oculta y usa el identificador de un modelo habilitado para Responses API
y salidas estructuradas en tu proyecto de OpenAI.

```bash
read -rs 'OPENAI_API_KEY?Clave API: '
export OPENAI_API_KEY
echo
read 'MODELO_ODS?Modelo de OpenAI: '
python analizar_ods.py \
  --seleccion resultados_inventario/20261118T205716004269Z/seleccion_analisis_verificada.json \
  --salida resultados_ods/historico --limite 25 --clasificar \
  --modelo "$MODELO_ODS" \
  --metas "config_local/Catalogo_ODS(5).xlsx" \
  --hoja "Metas ODS" --columna-codigo "Código Meta" \
  --columna-texto "Meta oficial Agenda 2030" \
  --tipos "config_local/BD_Actividades_ODSTEC(8).xlsx" \
  --directorio "config_local/Funcionarios_Directorio_TEC(1).xlsx"
```

Revisa un boletín con varias actividades, un anuncio, una actividad
realizada, un correo irrelevante y un afiche que aportó datos por OCR.
Contrasta la fecha de la actividad, la dependencia, el estado y las
metas sugeridas con las citas de la hoja Fuentes y los originales.

Para procesar los 933 correos, repite el comando sin `--limite 25` y
conserva la ruta de salida para reutilizar el avance. La base SQLite
actualiza por ID de Outlook y almacena el resultado de cada correo.
La extracción de adjuntos se reutiliza por SHA-256 y formato. Vuelve a
ejecutar el comando para recuperar errores de extracción o clasificación.
Evita ejecutar dos instancias al mismo tiempo sobre la misma salida.

El Excel contiene Actividades, Fuentes, Metas sugeridas, Control y Adjuntos.
La consolidación automática solo agrupa actividades con el mismo nombre,
fecha y dependencia normalizados. Las coincidencias parciales se revisan
manualmente. Los enlaces permanecen como referencias. Los videos, archivos
binarios sin formato identificado y GIF posteriores al primer fotograma
requieren revisión si son la única evidencia. El PDF se revisa página por
página y se aplica OCR a las páginas sin texto.

## Operación incremental

Mantén la carpeta local sincronizada con la salida de Power Automate.
Cuando termine el flujo y estén disponibles los tres directorios, cambia
`--seleccion …` por `--raiz "/ruta/de/la/extraccion" --inicio AAAA-MM-DD
--fin AAAA-MM-DD`. El script inventaría el rango y vuelve a usar la misma
base de avance, por lo que los correos ya clasificados con la misma
entrada y configuración se saltan. Los informes diarios y pendientes
quedan en la salida del inventario. Si se omite `--fin`, se usa el día
local de ejecución.

Para calendarizarlo en el Mac, configura el lanzador que utilices con
la ruta absoluta al Python de `.venv`, el directorio de trabajo y el
comando incremental, y programa su inicio después de que la extracción
esté disponible. Proporciona la clave por el mecanismo de secretos de
la máquina. Ejecuta dos veces el mismo rango para verificar que el
segundo pase no duplica correos ni actividades.

Usa `--solo-exportar --salida resultados_ods/historico` para reconstruir
el Excel a partir del avance guardado. Mantén los resultados, los
catálogos internos, los correos y las claves fuera del repositorio.
