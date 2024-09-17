# EventPro Organizer

**EventPro Organizer** es una aplicación de Google Apps Script diseñada para simplificar la gestión y planificación de eventos empresariales utilizando las herramientas de Google Workspace de forma integrada.

## Características

- Crear eventos automáticamente en Google Calendar.
- Generar formularios de registro personalizados con Google Forms.
- Organizar archivos relacionados con el evento en Google Drive.
- Enviar correos electrónicos de confirmación automáticos.
- Generar informes detallados de eventos y participantes.

## Requisitos previos

- Cuenta de Google Workspace (anteriormente G Suite).
- Permisos para crear y ejecutar Apps Scripts en tu cuenta.
- Acceso a Google Calendar, Drive, Forms, Gmail, y Spreadsheets.

## Instalación

1. Crea una nueva hoja de cálculo de Google.
2. Ve a `Extensiones > Apps Script`.
3. Copia y pega el contenido del archivo `Code.gs` en el editor de scripts.
4. Guarda el proyecto.

## Configuración

1. En el editor de scripts, ve a `Servicios` en el panel izquierdo.
2. Añade el servicio "People API".
3. Guarda los cambios.
4. Vuelve a la hoja de cálculo y actualiza la página.

## Uso

### Crear un nuevo evento

1. En la hoja de cálculo, haz clic en el menú `EventPro Organizer`.
2. Selecciona `Crear Nuevo Evento`.
3. Ingresa el nombre del evento cuando se te solicite.
4. El script creará automáticamente:
   - Un evento en tu calendario.
   - Una carpeta en Drive para los archivos del evento.
   - Un formulario de registro.
   - Un correo electrónico de confirmación.

### Generar informe de eventos

1. Haz clic en `EventPro Organizer > Generar Informe de Eventos`.
2. El script generará un informe con detalles de todos los eventos y lo guardará en tu Google Drive.

## Estructura del proyecto

- **`Code.gs`**: Contiene todo el código fuente del proyecto.
  - `onOpen()`: Crea el menú personalizado.
  - `crearNuevoEvento()`: Maneja la creación de nuevos eventos.
  - `generarInformeEventos()`: Genera el informe de eventos.

## Servicios de Google utilizados

- Google Calendar
- Google Drive
- Google Forms
- Gmail
- Google Sheets
- People API (Servicio avanzado)

## Solución de problemas

Si encuentras algún problema:

1. Verifica los logs de ejecución en el editor de scripts.
2. Asegúrate de tener los permisos necesarios para todos los servicios.
3. Comprueba que los IDs de los formularios y carpetas sean válidos y accesibles.

## Contribuir

Las contribuciones son bienvenidas. Por favor, abre un *issue* para discutir los cambios propuestos o crea un *pull request*.

## Licencia

Este proyecto está bajo la Licencia MIT. Consulta el archivo `LICENSE` para más detalles
