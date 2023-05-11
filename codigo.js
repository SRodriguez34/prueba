function enviarCorreo() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hoja 1");
    var startRow = 2; // La fila en la que empiezan los datos
    var numRows = sheet.getLastRow() - 1; // El número de filas a procesar
    var dataRange = sheet.getRange(startRow, 1, numRows, 7); // La cantidad de columnas del sheet
  
    var data = dataRange.getValues(); // Obtener los valores de las celdas
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var email = row[0]; // La dirección de correo electrónico
      var nombre = row[1]; // El nombre del destinatario
      var apellido = row[2]; // El apellido del destinatario
      var sendEmail = row[3]; // La casilla de verificación para enviar el correo electrónico
      var openTracking = row[4]; // La columna de registro de apertura del correo electrónico
      var enviado = row[5]; // La columna para el registro de envío del correo
      var rebotado = row[6]; // La columna para el registro de correos rebotados
  
      if (sendEmail == true && enviado != 'Sí') { // Si se marca la casilla de verificación y el correo no ha sido enviado previamente
        var asunto = "Asunto del correo";
        var mensaje = "Cuerpo del correo";
        var imagenUrl = "URL de la imagen de seguimiento";
  
        var template = HtmlService.createTemplateFromFile('correo_template');
        template.asunto = 'Este es el título del correo electrónico';
        template.saludo = 'Estimado';
        template.nombre = 'Juan';
        template.apellido = 'Pérez';
        template.cuerpo = 'Este es un párrafo de ejemplo en el que se puede incluir cualquier texto que se desee. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nulla pretium tortor id ante accumsan, vel sagittis nisi vestibulum. Ut faucibus libero id ex semper, non tincidunt elit venenatis.';
        template.imagen = 'https://acortar.link/HLzCho';
        var html = template.evaluate().getContent();
        var options = { htmlBody: html };
  
       
        try {
          GmailApp.sendEmail(email, asunto, 'Este es un mensaje', options);
          sheet.getRange(startRow + i, 6).setValue('Sí'); // Registrar que el correo fue enviado
        } catch (e) {
          sheet.getRange(startRow + i, 7).setValue('Sí'); // Registrar que el correo rebotó
        }
  
        Utilities.sleep(10000);
      }
    }
  }
  