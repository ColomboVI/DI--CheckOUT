function ThisSheet() {
  if (typeof ThisSheet.instancia === 'object') {
    return ThisSheet.instancia;
  }

  this._pais = new Pais(SpreadsheetApp.getActive().getRange('PAIS').getValue());
  this._programa = new Programa(SpreadsheetApp.getActive().getRange('PROGRAMA').getValue());
  this._directorio = null;
  this._directorio_workshops = null;
  this._directorio_checkout = null;
  /**************
     Feedback
  **************/
  this._directorio_feedback = null;

  var folders = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents();
  var encontrado = false;
  var folder;
  while (!encontrado && folders.hasNext()) {
    folder = folders.next();
    //Logger.log(folder.getName());
    encontrado = /^EVAL .+/.test(folder.getName());
  }
  if (encontrado) {
    this._directorio = folder;
    //como existe la carpeta, buscamos si ya estan generadas el resto de otros directorios donde se almancenan ficheros
    folders = this._directorio.getFolders();
    while (folders.hasNext()) {
      folder = folders.next();
      //Logger.log(folder.getName());
      // /^Respuestas FeedBack$/.test(folder.getName())
      if (/^Cuestionarios Workshops$/.test(folder.getName())) this._directorio_workshops = folder;
      else if (/^Test Check out/.test(folder.getName())) this._directorio_checkout = folder;
      /**************
         Feedback
      **************/
      else if (/^Respuestas FeedBack/.test(folder.getName())) this._directorio_feedback = folder;
    }
  }

  if (!this._directorio) throw 'No esta creado el directorio para esta hoja de evaluación';

  if (!this._directorio_workshops)
    this._directorio_workshops = this._directorio.createFolder('Cuestionarios Workshops');

  if (!this._directorio_checkout)
    this._directorio_checkout = this._directorio.createFolder('Test Check out');

  /**************
     Feedback
  **************/
  if (!this._directorio_feedback)
    this._directorio_feedback = this._directorio.createFolder('Respuestas FeedBack');

  //this._locale= new Locale(this._pais);

  this.addTestCheckout = function(formulario) {
    this._addFormulario(formulario, this._directorio_checkout);
     var htmlOutput = HtmlService
    .createHtmlOutput('<h3>Pulse el botón <em>"Enviar formulario"</em> cuando quiera enviar el formulario a los alumnos</h3>')
    .setWidth(650)
    .setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Formulario creado correctamente...');
  };
  this.addTestWorkshop = function(formulario) {
    this._addFormulario(formulario, this._directorio_workshops);
  };

  /**************
     Feedback
  **************/
  this.addTestFeedback = function(formulario) {
    this._addFormulario(formulario, this._directorio_feedback);
  };

  this.getRespuestasTestWorkshop = function(workshop) {
    var ficheros = this._directorio_workshops.getFiles();
    var fichero;
    var responses = [];
    var re = new RegExp('^Evaluación de Workshop de ' + workshop + ' \\(respuestas\\)$');
    while (ficheros.hasNext()) {
      fichero = ficheros.next();
      //Logger.log('fichero:'+fichero.getName()+' contra '+re.test(fichero.getName() ));
      if (re.test(fichero.getName())) responses.push(SpreadsheetApp.openById(fichero.getId()));
    }
    return responses;
  };

  this.getRespuestasTestCheckout = function() {
    return this._getCheckout(true);
  };
  this.getFormularioTestCheckout = function() {
    return this._getCheckout();
  };
  this._getCheckout = function(excel_responses) {
    var form = null,
        folder = this._directorio_checkout;

    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      if (
        (excel_responses && /Check-out \(respuestas\)$/.test(file.getName())) ||
        (!excel_responses && /Check-out$/.test(file.getName()))
        ) {
          form = file;
        }
    }
    return form;
  };

  /**************
     Feedback
  **************/
   this.getRespuestasTestFeedback = function() {
    return this._getFeedback(true);
  };
  this.getFormularioTestFeedback = function() {
    return this._getFeedback();
  };
    this._getFeedback = function(excel_responses) {
    var form = null,
        folder = this._directorio_feedback;

    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      if (
        (excel_responses && /Feedback \(respuestas\)$/.test(file.getName())) ||
        (!excel_responses && /Cuestionario Alumnos Datio Immersion$/.test(file.getName()))
        ) {
          form = file;
        }
    }
    return form;
  };

  this._addFormulario = function(formulario, carpeta) {
    var formulario_drive = DriveApp.getFileById(formulario.getFormulario().getId());
    carpeta.addFile(formulario_drive);

    var excel_drive = DriveApp.getFileById(formulario.getExcelAsociado().getId());
    carpeta.addFile(excel_drive);
  };

  this.getPrograma = function() {
    return this._programa;
  };

  this.getPais = function() {
    return this._pais;
  };
  this.getDirectorio = function() {
    return this._directorio;
  };

  ThisSheet.instancia = this;
  //return this;
}
