# Sistema de Combinación de Correspondencia con Google Apps Script

Este proyecto implementa un sistema de combinación de correspondencia utilizando **Google Apps Script**, **Google Sheets** y **Google Docs**. Permite generar automáticamente documentos personalizados en PDF a partir de una plantilla y datos en una hoja de cálculo, y enviarlos por correo electrónico.

## Escenario real de aplicación:
Imagina que eres el coordinador de un programa de becas en una universidad. Cada semestre, necesitas enviar cartas de aceptación personalizadas a cientos de estudiantes. Aquí es donde este script sería muy útil:

Crea una hoja de cálculo con columnas como: Nombre, Email, Carrera, Monto de la beca, Fecha de inicio, etc.
Diseña una plantilla de carta de aceptación en Word o Google Docs, usando marcadores de posición para los datos personalizados.
Ejecuta el script para generar automáticamente todas las cartas, convertirlas a PDF, y enviarlas por correo electrónico a cada estudiante.

Este proceso automatizado te ahorraría horas de trabajo manual, reduciría errores y proporcionaría una experiencia consistente y profesional para todos los estudiantes. Además, al tener los PDFs compartidos públicamente, los estudiantes podrían acceder a sus cartas de aceptación en cualquier momento, incluso si pierden el correo electrónico original.

## Características

- Genera un menú personalizado en Google Sheets para iniciar el proceso.
- Utiliza una plantilla de Google Docs para crear documentos personalizados.
- Convierte los documentos a PDF automáticamente.
- Comparte los PDFs públicamente y envía enlaces por correo electrónico.
- Procesa múltiples registros en una sola ejecución.

## Requisitos previos

- Una cuenta de Google con acceso a Google Drive, Google Sheets y Google Docs.
- Permisos para crear y ejecutar Apps Scripts en tu cuenta de Google.

## Configuración

### 1. Crear la hoja de cálculo de datos:

1. Crea una nueva hoja de cálculo en Google Sheets.
2. Añade los datos de los destinatarios con encabezados en la primera fila.
3. Asegúrate de incluir una columna para el correo electrónico.

### 2. Crear la plantilla del documento:

1. Crea un nuevo documento en Google Docs.
2. Diseña la plantilla usando marcadores de posición como `{{Nombre}}`, `{{Email}}`, etc.
3. Los marcadores deben coincidir con los encabezados de tu hoja de cálculo.

### 3. Crear una carpeta para los PDFs:

1. Crea una nueva carpeta en Google Drive para almacenar los PDFs generados.

### 4. Obtener los IDs necesarios:

- **Para la plantilla**: Abre el documento y copia el ID de la URL (la parte entre `/d/` y `/edit`).
- **Para la carpeta de PDFs**: Abre la carpeta y copia el ID de la URL (la parte después de `/folders/`).

### 5. Configurar el script:

1. En la hoja de cálculo, ve a **Extensiones > Apps Script**.
2. Copia y pega el código del script proporcionado.
3. Reemplaza los valores de `templateId` y `folderId` con los IDs que obtuviste.

## Uso

1. Abre la hoja de cálculo de Google Sheets que contiene los datos.
2. Verás un nuevo menú llamado **"Combinar Correspondencia"**.
3. Selecciona **"Combinar Correspondencia" > "Iniciar Combinación"**.
4. Concede los permisos necesarios cuando se te soliciten.
5. El script generará los documentos, los convertirá a PDF y enviará los correos electrónicos.

## Personalización

- Modifica la plantilla de Google Docs según tus necesidades.
- Ajusta el contenido del correo electrónico en la función `iniciarCombinacion()` del script.
- Añade o modifica campos en la hoja de cálculo y actualiza la plantilla en consecuencia.

## Solución de problemas

- Verifica que los IDs de la plantilla y la carpeta de PDFs sean correctos.
- La primera columna debe ser el campo Nombre

## Contribuciones

Las contribuciones a este proyecto son bienvenidas. Por favor, abre un *issue* o un *pull request* para sugerir cambios o mejoras.

## Licencia

Este proyecto está bajo la licencia MIT. Consulta el archivo LICENSE para más detalles.
