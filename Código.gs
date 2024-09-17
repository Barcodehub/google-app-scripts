function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Combinar Correspondencia')
    .addItem('Iniciar Combinación', 'iniciarCombinacion')
    .addToUi();
}

function iniciarCombinacion() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var templateId = '1DMNW6IIkjDMfvFu4R_OITmjGXXT9Pkf2QA8PGvADMDQ'; // ID del documento de Word plantilla
  var folderId = '1iTS7EgN2v0ZwtQprF26Frj8Qxny1Oiyx'; // ID de la carpeta donde se guardarán los PDFs

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var docCopy = DriveApp.getFileById(templateId).makeCopy();
    var doc = DocumentApp.openById(docCopy.getId());
    var body = doc.getBody();

    // Reemplazar placeholders en el documento
    for (var j = 0; j < headers.length; j++) {
      body.replaceText('{{' + headers[j] + '}}', row[j]);
    }

    doc.saveAndClose();

    // Convertir a PDF
    var pdf = DriveApp.getFileById(docCopy.getId()).getAs('application/pdf');
    var pdfFile = DriveApp.createFile(pdf);
    pdfFile.setName(row[0] + ' - Documento.pdf'); // Asumir q Nombre esta en la primera columna

    // Mover a la carpeta y compartir publicamente
    DriveApp.getFolderById(folderId).addFile(pdfFile);
    DriveApp.removeFile(pdfFile);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Enviar email
    var emailAddress = row[headers.indexOf('Email')];
    var subject = 'Su documento personalizado';
    var body = 'Adjunto encontrará su documento personalizado. Puede acceder al PDF en el siguiente enlace: ' + pdfFile.getUrl();
    MailApp.sendEmail(emailAddress, subject, body);

    // Limpiar: eliminar la copia del documento
    DriveApp.getFileById(docCopy.getId()).setTrashed(true);
  }

  SpreadsheetApp.getUi().alert('Combinación de correspondencia completada.');
}