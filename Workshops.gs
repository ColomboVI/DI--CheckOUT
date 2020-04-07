

function leerRespuestasWorkshop(workshop)
{
      var sheet_respuestas=SpreadsheetApp.getActive().getSheetByName('WS '+workshop+' - Rubricas (TEST)');
      
      var excel_response=(new ThisSheet()).getRespuestasTestWorkshop(workshop),
          last_row_sheet_respuestas=1,
          last_column_excel_response,
          last_row_excel_response;
      
      for(var i=0;i<excel_response.length;i++)
      {
        //el primero se copiara entero incluido cabeceras
        //el resto desde la segunda fila
        last_column_excel_response=excel_response[i].getActiveSheet().getLastColumn();
        last_row_excel_response=excel_response[i].getActiveSheet().getLastRow();
          //Logger.log(JSON.stringify(excel_response[i].getActiveSheet().getRange((encontrado)?2:1, 2, last_row_excel_response-((encontrado)?1:0), last_column_excel_response-1).getValues()));
          //Logger.log(JSON.stringify(sheet_respuestas.getRange((encontrado)?last_row_sheet_respuestas:1, 1, 
          //                                                    last_row_excel_response-((encontrado)?1:0), last_column_excel_response-1).getValues()));
          sheet_respuestas.getRange(last_row_sheet_respuestas, 1, 
                                    last_row_excel_response-((i)?1:0), last_column_excel_response-1)
          .setValues(excel_response[i].getActiveSheet().getRange((i)?2:1, 2, last_row_excel_response-((i)?1:0), last_column_excel_response-1).getValues())
          
          last_row_sheet_respuestas += last_row_excel_response-((i)?1:0);
      }
      SpreadsheetApp.flush();
}


function lanzarCuestionarioWorkshop(workshop)
{
  var evaluados=getEmailsSelected(true).join(';');
  var test=TestWorkshop(workshop,evaluados);
  (new ThisSheet()).addTestWorkshop(test);
  DriveApp.getRootFolder().removeFile(DriveApp.getFileById(test.getFormulario().getId()));
  DriveApp.getRootFolder().removeFile(DriveApp.getFileById(test.getExcelAsociado().getId()));
  (new Comunicaciones()).mandarEvaluacionWorkshop(getEmailsSelected(false),evaluados, test.getFormulario(),workshop);
}


function getEmailsSelected(selected)
{
   var sheet_workshops=SpreadsheetApp.getActive().getSheetByName('Workshops');
   var max_registros=sheet_workshops.getLastRow();
   var table_emails=sheet_workshops.getRange(4,1,max_registros, 3).getValues();
   var emails=[];
   //Logger.log(JSON.stringify(table_emails))

   for (var i=0; i<max_registros && table_emails[i][1];i++)
   {
       if (table_emails[i][0]==selected && table_emails[i][2]!='Baja')
           emails.push(table_emails[i][1].trim())          
   }
   //Logger.log(JSON.stringify(emails))
   
   return emails;
}
