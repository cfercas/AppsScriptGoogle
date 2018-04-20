//se ejecuta si se pulsa la opcion del menu
function principal(){
                //hoja de calculo activa
  var ss = SpreadsheetApp.getActiveSpreadsheet();
                //la primera hoja
  var sheet = ss.getSheets()[0];
                //recoge los datos
  var range = sheet.getDataRange();
                //cuenta las filas
  var numRows = range.getNumRows() ;
                //inicializa la variable m
  var m = 0;
                //bucle que recorre todas las filas
  for(var i = 2; i <= numRows; i++) {
                  //valor de la celda segun la columna 2
    var proceso = range.getCell(i, 2);
                  //valor de la celda segun la columna 1
    var estado = range.getCell(i, 1);
                  //declarar variables
    var datos = new Array();
                  //comprueba si el estado y el proceso de las trasacciones
    if((proceso.getValue() != "Enviado") && (estado.getValue() == "Aceptado")){
                    /*Logger.log(estado.getValue());
                    Logger.log(proceso.getValue());*/
                    //guardamos los valores de los campos que necesitamos en
      var fecha = (range.getCell(i, 4)).getValue();
      var nombre = (range.getCell(i, 5)).getValue();
      var concepto = (range.getCell(i, 6)).getValue();
      var email = (range.getCell(i, 9)).getValue();
      var responsable = (range.getCell(i, 3)).getValue();
      var cuenta = (range.getCell(i, 8)).getValue();
      var importe = (range.getCell(i, 7)).getValue();
      var origen = (range.getCell(i, 10)).getValue();
      var destino = (range.getCell(i, 11)).getValue();

                    // Subject
      var subject = "Estado de solicitud de transacciones Aceptada por " + responsable;

                    // emailBody
      var emailBody = "Usuario que realizó la petición: " + nombre +
                    "\nFecha:  " + fecha +
                    "\nEmail:  " + email +
                    "\nConcepto:  " + concepto +
                    "\nCuenta:  " + cuenta +
                    "\nWith importe:  " + importe +
					"\nRegister on " + fecha +
                    "\n\nGracias y un Saludo!";

                    // html
      var htmlBody =  "Solicitud de transacción enviada en fecha:" + fecha + "</i>" +
					"<br/><br/>The details you entered were as follows: " +
					"<br/>usuarios: <font color=\"red\"><strong>" + nombre + "</strong></font>" +
					"<br/>Concepto: " + concepto +
					"<br/>Importe: " + importe;

                    // More info for Advanced Options Parameters
                    // https://developers.google.com/apps-script/reference/mail/mail-app#sendEmail(String,String,String,Object)
      var advancedOpts = { name: "Solicitudes Aceptadas", htmlBody: htmlBody };

                    // Envio de email
                    //Logger.log(subject);
      MailApp.sendEmail("...@GMAIL.COM", subject, emailBody, advancedOpts);

      range.getCell(i, 2).setValue("Enviado");
    }
  }
}
              //crear un menu para llamar a las funciones que creemos
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu')
      .addItem('Enviar transacciones aceptadas', 'menuItem1')
      .addSeparator()
      .addSubMenu(ui.createMenu('Sub-menu')
          .addItem('Segundo', 'menuItem2'))
      .addToUi();

}

              //opciones del menu
function menuItem1() {
  var ui = SpreadsheetApp.getUi() // Puede ser un documento o un Formulario
  var preguntar = ui.alert('Estas seguro que quieres enviar las transacciones???', ui.ButtonSet.YES_NO);
                //si la opcion a la alerta es SI entonces ejecuta la función
  if (preguntar == ui.Button.YES){
    principal();
  }

}


function menuItem2() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .alert('Aun no hay vinculada ninguna función a este items!!');
}
//Al rellenar un formulario se generará esta función, eso se puede configurar en el menu 'Editar' -> 'Activadores del proyecto activo'
function emailOnFormSubmit() {
	              //Variables para nuestra  hoja de calculo
                  //hoja de calculo activa
    var ss = SpreadsheetApp.getActiveSpreadsheet();
                  //la primera hoja
    var sheet = ss.getSheets()[0];
                 //recoge los datos
    var range = sheet.getDataRange();
                 //recoge la ultima fila
    var ultimo = sheet.getLastRow();
                  //var ordenar = sheet.getDataRange() + ultimo;
                  //ordenar.sort(1);
  //Las variables
      var fecha = (range.getCell(ultimo, 4)).getValue();
      var nombre = (range.getCell(ultimo, 5)).getValue();
      var concepto = (range.getCell(ultimo, 6)).getValue();
      var correo = (range.getCell(ultimo, 9)).getValue();
      var responsable = (range.getCell(ultimo, 3)).getValue();
      var cuenta = (range.getCell(ultimo, 8)).getValue();
      var importe = (range.getCell(ultimo, 7)).getValue();
      var origen = (range.getCell(ultimo, 10)).getValue();
      var destino = (range.getCell(ultimo, 11)).getValue();

                    //Logger.log(nombre);
                    //variable que guardará el email del responsable del fichero
    var email = "";
                  //enlace adjunto al correo para que puedan poner el tipo de Estado si se acepta o no
                  //var url = "https://docs.google...";
                  //opciones de responsable aqui se añadirán aquellos usuarios a los que se les enviará la notificación de una nueva transacción
    switch (responsable)  {
        case "Responsable 1":
            email = "...";
                         // Logger.log(email);
            //lo guarda en un hoja nueva del excel que despues importare en otra con IMPORTRANGE
            var sheet1 = ss.getSheets()[1];
            //el ultimo de esta hoja
            var ultimo1 = sheet1.getLastRow();
            //variable para espeficcar el rango donde guardar los nuevos datos
            var fila = 1;
            for(var i = 1; i <= ultimo1; i++) {
              fila = fila + 1;
            };
            //guardamos los valores en un array
            var values = [
              [ "","", responsable, fecha, nombre, concepto, importe, cuenta, correo, origen, destino]
            ];
            //especficamos el rango donde vamos a guardar los datos
            var range1 = sheet1.getRange("A"+fila+":K"+fila);
            //los guardamos
            range1.setValues(values);
            //Logger.log(values); //nos sirve para verificar el valor de las variables
            var url = "https://docs.google.com/spreadsheets/d/... url que le llega al responsable con un enlace a el excel al que tiene permisos";
            //teminamos el bucle
            break;

        case "Responsable 2":
            email = "...@gmail.com";
                         //Logger.log(email);
            var sheet2 = ss.getSheets()[2];
            var ultimo2 = sheet2.getLastRow();
            var fila = 1;
            for(var i = 1; i <= ultimo2; i++) {
              fila = fila + 1;
            };
            var values = [
              [ "","", responsable, fecha, nombre, concepto, importe, cuenta, correo, origen, destino]
            ];
            var range2 = sheet2.getRange("A"+fila+":K"+fila);
            range2.setValues(values);
            Logger.log(values);
            var url = "https://docs.google.com/spreadsheets....";
            break;
        case "Responsable 3":
            email = "...";
            var sheet3 = ss.getSheets()[3];
            var ultimo3 = sheet3.getLastRow();
            var fila = 1;
            for(var i = 1; i <= ultimo3; i++) {
              fila = fila + 1;
            };
            var values = [
              [ "","", responsable, fecha, nombre, concepto, importe, cuenta, correo, origen, destino]
            ];
            var range3 = sheet3.getRange("A"+fila+":K"+fila);
            range3.setValues(values);
            //Logger.log(values);
            var url = "...";
            break;
        case "Responsable 4":
            email = "...";
            var sheet4 = ss.getSheets()[4];
            var ultimo4 = sheet4.getLastRow();
            var fila = 1;
            for(var i = 1; i <= ultimo4; i++) {
              fila = fila + 1;
            };
            var values = [
              [ "","", responsable, fecha, nombre, concepto, importe, cuenta, correo, origen, destino]
            ];
            var range4 = sheet4.getRange("A"+fila+":K"+fila);
            range4.setValues(values);
            Logger.log(values);
            var url = "https://docs.google.com/spreadsheets/d/...";
            break;
        default:
            email = "...";
            var sheet5 = ss.getSheets()[5];
            var ultimo5 = sheet5.getLastRow();
            var fila = 1;
            for(var i = 1; i <= ultimo5; i++) {
              fila = fila + 1;
            };
            var values = [
              [ "","", responsable, fecha, nombre, concepto, importe, cuenta, correo, origen, destino]
            ];
            var range5 = sheet5.getRange("A"+fila+":K"+fila);
            range5.setValues(values);
            Logger.log(values);
            var url = "https://docs.google.com/spreadsheets/...";
            break;
    }

	                 // Subject de nuestro envio de correo
  var subject = "Nueva Solicitud de transacciones de: " + nombre;

	                // emailBody es par alos dispositivos que no pueden renderizar html
  var emailBody = "Solicitud de transacción enviada por: " + nombre +
                    "\nConcepto:  " + concepto +
                    "\nWith importe:  " + importe +
					"\nRegister on " + fecha +
                    "\n\nPara poder atender la petición accede a: "+ url +" !";

	               // html is for those devices that can render HTML
	               // nowadays almost all devices can render HTML
	var htmlBody =  "Solicitud de transacción enviada en fecha:" + fecha + "</i>" +
					"<br/><br/>The details you entered were as follows: " +
					"<br/>usuarios: <font color=\"red\"><strong>" + nombre + "</strong></font>" +
					"<br/>Concepto: " + concepto +
					"<br/>Importe: " + importe +
                    "<br/>Enlace: " + url;

	              // More info for Advanced Options Parameters
	              // https://developers.google.com/apps-script/reference/mail/mail-app#sendEmail(String,String,String,Object)
	var advancedOpts = { name: "Formulario - Solicitudes", htmlBody: htmlBody };

	              // Envio de email
                  //Logger.log(subject);
	MailApp.sendEmail(email, subject, emailBody, advancedOpts);

}
