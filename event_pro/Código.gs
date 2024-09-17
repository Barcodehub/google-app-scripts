// Archivo: Code.gs

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('EventPro Organizer')
    .addItem('Crear Nuevo Evento', 'crearNuevoEvento')
    .addItem('Generar Informe de Eventos', 'generarInformeEventos')
    .addToUi();
}

function crearNuevoEvento() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Crear Nuevo Evento', 'Ingrese el nombre del evento:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var nombreEvento = response.getResponseText();
    var fechaEvento = new Date(new Date().getTime() + (7 * 24 * 60 * 60 * 1000)); // Una semana desde hoy
    
    // Crear evento en el calendario
    var calendario = CalendarApp.getDefaultCalendar();
    var evento = calendario.createEvent(nombreEvento, fechaEvento, new Date(fechaEvento.getTime() + (2 * 60 * 60 * 1000)));
    evento.setVisibility(CalendarApp.Visibility.PUBLIC); // Hacer el evento público
    
    // Crear carpeta en Drive
    var folder = DriveApp.createFolder('Evento: ' + nombreEvento);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // Compartir la carpeta
    
    // Crear formulario de registro
    var form = FormApp.create('Registro para ' + nombreEvento);
    form.setAllowResponseEdits(true)
        .setAcceptingResponses(true)
        .setRequireLogin(false); // Permitir respuestas sin iniciar sesión
    form.addTextItem().setTitle('Nombre');
    form.addTextItem().setTitle('Email');
    DriveApp.getFileById(form.getId()).moveTo(folder);
    DriveApp.getFileById(form.getId()).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT); // Compartir el formulario
    
    // Enviar email de confirmación
    GmailApp.sendEmail(Session.getActiveUser().getEmail(), 'Nuevo evento creado: ' + nombreEvento,
      'Se ha creado un nuevo evento: ' + nombreEvento + '\n' +
      'Fecha: ' + fechaEvento.toDateString() + '\n' +
      'Formulario de registro: ' + form.getPublishedUrl() + '\n' +
      'Carpeta del evento: ' + folder.getUrl());
    
    // Guardar información en la hoja de cálculo
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.appendRow([nombreEvento, fechaEvento, evento.getId(), folder.getId(), form.getId(), form.getPublishedUrl()]);
    
    ui.alert('Evento creado con éxito. El formulario de registro está disponible en: ' + form.getPublishedUrl());
  }
}

function generarInformeEventos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var informe = "Informe de Eventos\n\n";
  
  console.log("Número total de filas: " + data.length); // Depuración
  
  for (var i = 1; i < data.length; i++) {
    console.log("Procesando fila: " + i); // Depuración
    
    var nombreEvento = data[i][0];
    var fechaEvento = new Date(data[i][1]);
    var formId = data[i][4];
    
    console.log("Nombre del evento: " + nombreEvento); // Depuración
    console.log("Fecha del evento: " + fechaEvento); // Depuración
    console.log("ID del formulario: " + formId); // Depuración
    
    informe += "Evento: " + nombreEvento + "\n";
    informe += "Fecha: " + fechaEvento.toDateString() + "\n";
    
    // Obtener respuestas del formulario
    try {
      var form = FormApp.openById(formId);
      var respuestas = form.getResponses();
      informe += "Registrados: " + respuestas.length + "\n";
      
      // Usar People API para obtener más detalles de los participantes
      if (respuestas.length > 0) {
        var email = respuestas[0].getResponseForItem(form.getItems()[1]).getResponse();
        var person = People.People.get('people/me', {personFields: 'names,emailAddresses'});
        if (person.names && person.names.length > 0) {
          informe += "Ejemplo de participante: " + person.names[0].displayName + "\n";
        }
      }
    } catch (error) {
      console.error("Error al procesar el formulario: " + error); // Depuración
      informe += "Error al procesar las respuestas del formulario.\n";
    }
    
    informe += "\n";
  }
  
  // Guardar el informe en Drive
  var reportFile = DriveApp.createFile('Informe de Eventos.txt', informe);
  
  // Guardar la última fecha de generación del informe
  PropertiesService.getScriptProperties().setProperty('ultimoInforme', new Date().toISOString());
  
  console.log("Informe generado con ID: " + reportFile.getId()); // Depuración
  SpreadsheetApp.getUi().alert('Informe generado y guardado en Drive con ID: ' + reportFile.getId());
}