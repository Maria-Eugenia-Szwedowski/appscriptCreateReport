📌 Generador de Informes desde Google Sheets a Google Docs y PDF

Este proyecto fue desarrollado por María Eugenia Szwedowski con el objetivo de automatizar la generación de informes a partir de datos almacenados en Google Sheets, utilizando Google Apps Script.

🛠️ Funcionalidad Principal

✔️ Toma datos de una hoja de cálculo en Google Sheets.
✔️ Usa un documento de plantilla en Google Docs.
✔️ Rellena la plantilla con los datos de cada fila.
✔️ Genera automáticamente un archivo .docx y un .pdf.
✔️ Guarda los archivos en una carpeta específica en Google Drive.
📂 Organización del Código
El código está estructurado en un solo archivo.


📎 Requisitos

Tener acceso a Google Sheets, Google Docs y Google Drive.
Crear una plantilla en Google Docs con los marcadores <<nombre>>, <<fecha>>, etc.
Configurar los IDs de la plantilla y la carpeta de destino


📌 Paso a Paso

1️⃣ Definición de variables clave

TEMPLATE_FILE_ID: ID del archivo de Google Docs que se usará como plantilla.
DESTINATION_FOLDER_ID: ID de la carpeta de Google Drive donde se guardarán los archivos generados.
sheet.getActiveSheet(): Obtiene la hoja activa.
sheet.getDataRange().getValues(): Extrae todos los datos de la hoja en una matriz (formValues).
docTemplate: Abre el documento de plantilla.
folder: Accede a la carpeta donde se guardarán los archivos.
headers: Guarda la primera fila de la hoja de cálculo (encabezados de las columnas).
Se empieza desde rowIndex = 1 para omitir la fila de encabezados.
formRow: Guarda los datos de la fila actual.
fileName: Obtiene un valor único (columna 3) para nombrar el archivo.
Se crea una copia de la plantilla.
copyDoc abre la copia como un documento editable.


📌 Resumen del Flujo

✅ Lee los datos de Google Sheets
✅ Abre la plantilla en Google Docs
✅ Crea una copia del documento
✅ Reemplaza los marcadores con los datos
✅ Guarda los archivos generados en la carpeta de Drive
✅ Convierte el documento a PDF
✅ Guarda el PDF en la carpeta de destino
