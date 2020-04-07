function onOpen() {
  var ui = SpreadsheetApp.getUi();
    ui.createMenu('Evaluacion general curso')
      .addItem('Crear cuestionario alumnos', 'crearFormulario')
      .addSeparator()
      .addItem('Enviar cuestionario alumnos', 'enviarFormularioFeedback')
      .addSeparator()
      .addItem('Recopilar respuestas', 'recopilarRespuestasFeedback')
      .addToUi();      
}
/*******************************************************************
/ 1.) Coger emails (DONE)
/ Coger emails de la hoja de Desarrollo (solo si no estan de baja).
/ devuelve un array de emails.
********************************************************************/
function getEmailsEncuesta() {
  var sheet_workshops = SpreadsheetApp.getActive().getSheetByName('Desarrollo');
  var max_registros = sheet_workshops.getLastRow();
  var table_emails = sheet_workshops.getRange(3, 1, max_registros, 3).getValues();
  Logger.log(table_emails)
  var emails = [];
  for (var i = 0; i < max_registros && table_emails[i][1]; i++) {
    if (table_emails[i][2] != 'Baja') {
      emails.push(table_emails[i][1].trim());
    }
  }
  return emails;
}
/**************************************************** 
 * 2.) Crear una carpeta destino de respuestas (DONE) *
 ****************************************************/ 
 function crearCarpeta(){
 var folders = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents();
 var carpetaRespuestas = null;
 var encontrado = false;
 var folder; 
 while (!encontrado && folders.hasNext()) {
    folder = folders.next();
    //Logger.log(folder.getName());
    encontrado=/^EVAL .+/.test(folder.getName())
  }
   if(encontrado) {
       directorio=folder;
       //como existe la carpeta, buscamos si ya estan generadas el resto de otros directorios donde se almancenan ficheros
       folders = directorio.getFolders();
       while (folders.hasNext())  {
         folder = folders.next();
         //Logger.log(folder.getName());
         if(/^Respuestas FeedBack$/.test(folder.getName())){
         //Logger.log("Entra aqui")
         return carpetaRespuestas = folder;          
         }else{
         //Logger.log("Entra alla")
         return directorio.createFolder('Respuestas FeedBack');         
         }
      }
    }
  }
 /**************************************************** 
 * 3.) Crear formulario y enviar (DONE) *
 ****************************************************/
function crearFormulario() {
  /************************************************************************************
 *                   Plantilla de formulario de DI                                   * 
 * https://docs.google.com/forms/d/1soLhHnCuPgjvVqkgCVGB6mN7LkkO9PhgJdn8HvZeoSE/edit *
 *************************************************************************************/  
   var formulario = DriveApp.getFileById('1soLhHnCuPgjvVqkgCVGB6mN7LkkO9PhgJdn8HvZeoSE').makeCopy('Cuestionario Alumnos Datio Immersion',DriveApp.getRootFolder()); 
   var form = FormApp.openById(formulario.getId());  
   var carpeta = crearCarpeta()
   var formulario_drive= DriveApp.getFileById(form.getId()); 
          carpeta.addFile(formulario_drive);       
          // Comprobar que si existe formulario y si no crearlo
          // crearForm()                    
          var excelRespuestas = SpreadsheetApp.create('Feedback (respuestas)');
          //enlazamos con formulario
          //Logger.log(excelRespuestas.getId())              
          form.setDestination(FormApp.DestinationType.SPREADSHEET, excelRespuestas.getId());
          var respuestasExcell = DriveApp.getFileById(excelRespuestas.getId());
          carpeta.addFile(respuestasExcell);         
          Logger.log('generarCopiarTest] para '+ '' +' con id '+excelRespuestas.getId()+' es destino del formulario '+formulario_drive.getName())  
          //new Comunicaciones().mandarfeedback(getEmailsEncuesta(), form)
  }  
  function enviarFormularioFeedback(){
       var this_sheet = new ThisSheet();
       var encuesta = this_sheet.getFormularioTestFeedback();
  
      if(!encuesta){
           var htmlOutput = HtmlService
        .createHtmlOutput('<h3>Debe crear un formulario antes de enviar</h3>')
        .setWidth(650)
        .setHeight(200);
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'No existe ningún formulario creado');
      } else if (encuesta){
          new Comunicaciones().mandarfeedback(getEmailsEncuesta(), FormApp.openById(encuesta.getId()));
       }  
}
   function recopilarRespuestasFeedback() {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Encuesta');                     
        var excel_response=(new ThisSheet()).getRespuestasTestFeedback();
        //cogermos el valor en la celda A1
        //si no es vacio entocnes abrimos la hoja y copiamos los valores
        if (excel_response) {
          var excel_response_ss= SpreadsheetApp.open(excel_response);          
          var last_column_excel_response = excel_response_ss.getLastColumn();
          if (excel_response_ss.getLastRow()-1>0) {
           sheet.getRange(2, 3, excel_response_ss.getLastRow(), excel_response_ss.getLastColumn()).setValues(excel_response_ss.getActiveSheet().getRange(2, 1, excel_response_ss.getLastRow(), excel_response_ss.getLastColumn()).getValues())
          }
        }
      } 
