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
    
    try {
      // Crear evento en el calendario
      var calendario = CalendarApp.getDefaultCalendar();
      var evento = calendario.createEvent(nombreEvento, fechaEvento, new Date(fechaEvento.getTime() + (2 * 60 * 60 * 1000)));
      evento.setVisibility(CalendarApp.Visibility.PUBLIC);
      console.log("Evento creado en el calendario: " + evento.getId());
      
      // Crear carpeta en Drive
      var folder = DriveApp.createFolder('Evento: ' + nombreEvento);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      console.log("Carpeta creada en Drive: " + folder.getId());
      
      // Crear formulario de registro
      var form = FormApp.create('Registro para ' + nombreEvento);
      form.setAllowResponseEdits(true)
          .setAcceptingResponses(true)
          .setRequireLogin(false);
      form.addTextItem().setTitle('Nombre');
      form.addTextItem().setTitle('Email');
      DriveApp.getFileById(form.getId()).moveTo(folder);
      DriveApp.getFileById(form.getId()).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
      console.log("Formulario creado: " + form.getId() + ", URL: " + form.getPublishedUrl());
      
      // Enviar email de confirmación
      var emailBody = 'Se ha creado un nuevo evento: ' + nombreEvento + '\n' +
                      'Fecha: ' + fechaEvento.toDateString() + '\n' +
                      'Formulario de registro: ' + form.getPublishedUrl() + '\n' +
                      'Carpeta del evento: ' + folder.getUrl();
      GmailApp.sendEmail(Session.getActiveUser().getEmail(), 'Nuevo evento creado: ' + nombreEvento, emailBody);
      console.log("Correo enviado a: " + Session.getActiveUser().getEmail());
      
      // Guardar información en la hoja de cálculo
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      sheet.appendRow([nombreEvento, fechaEvento, evento.getId(), folder.getId(), form.getId(), form.getPublishedUrl()]);
      console.log("Información guardada en la hoja de cálculo");
      
      ui.alert('Evento creado con éxito. El formulario de registro está disponible en: ' + form.getPublishedUrl());
    } catch (error) {
      console.error("Error al crear el evento: " + error.toString());
      ui.alert('Error al crear el evento: ' + error.toString());
    }
  }
}

function generarInformeEventos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var informe = "Informe de Eventos\n\n";
  
  console.log("Generando informe para " + (data.length - 1) + " eventos");
  
  for (var i = 1; i < data.length; i++) {
    try {
      console.log("Procesando evento #" + i);
      
      var nombreEvento = data[i][0];
      var fechaEvento = new Date(data[i][1]);
      var formId = data[i][4];
      
      informe += "Evento: " + nombreEvento + "\n";
      informe += "Fecha: " + fechaEvento.toDateString() + "\n";
      
      // Obtener respuestas del formulario
      var form = FormApp.openById(formId);
      var respuestas = form.getResponses();
      informe += "Registrados: " + respuestas.length + "\n";
      
      console.log("Formulario " + formId + " tiene " + respuestas.length + " respuestas");
      
      // Comentar o eliminar la sección de People API
      /*
      // Usar People API para obtener más detalles de los participantes
      if (respuestas.length > 0) {
        var email = respuestas[0].getResponseForItem(form.getItems()[1]).getResponse();
        var person = People.People.get('people/me', {personFields: 'names,emailAddresses'});
        if (person.names && person.names.length > 0) {
          informe += "Ejemplo de participante: " + person.names[0].displayName + "\n";
        }
      }
      */
    } catch (error) {
      console.error("Error al procesar el evento #" + i + ": " + error.toString());
      informe += "Error al procesar las respuestas del formulario: " + error.toString() + "\n";
    }
    
    informe += "\n";
  }
  
  // Guardar el informe en Drive
  var reportFile = DriveApp.createFile('Informe de Eventos.txt', informe);
  console.log("Informe guardado en Drive con ID: " + reportFile.getId());
  
  // Guardar la última fecha de generación del informe
  PropertiesService.getScriptProperties().setProperty('ultimoInforme', new Date().toISOString());
  
  SpreadsheetApp.getUi().alert('Informe generado y guardado en Drive con ID: ' + reportFile.getId());
}