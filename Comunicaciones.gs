/*** LISTADO DE FUNCIONES de la clase Comunicaciones:
 ------------------------------
 + mandarCorreoCheckOut
 + mandarEvaluacionWorkshop
***/

function Comunicaciones() {
  if (typeof Comunicaciones.instancia === 'object') {
    return Comunicaciones.instancia;
  }

  Comunicaciones.instancia = this;
  //imagen cabecera del correo
  this._postal_blob = DriveApp.getFileById('1c5icGsDz0ge_o-fNJZw6HP0R0s6tOIUK').getBlob().setName("postal");

  //obtenemos la referencia a la clase que tiene los mensajes que van a ir en el correo
  this._mensajes = Mensajes.getInstancia();
  //obtenemos la referencia a la clase que maneja los paises y los buzones
  this._pais = (new ThisSheet()).getPais();

  //mandar un correo con remite una direccion de correo de buzon de pais REQUIERE TENERLO CONFIGURADO EN LA CUNTA PERSONAL COMO ALIAS
  this._enviarMailDesdeBuzonPais = function (to, options, text) {

    //Logger.log('_enviarMail] empieza mandarMail a '+to+' sobre: '+text)

    options.name = this._pais.getNombreEmisor();
    options.from = this._pais.getDireccionBuzon();

    GmailApp.sendEmail(to, this._mensajes.getAsuntoEmail(), text, options);

  }
  //mandar un correo con remite una direccion de correo personal
  this._enviarMailDesdeBuzonPersonal = function (to, asunto, options, text) {

    GmailApp.sendEmail(to, asunto, text, options);

  }

  //dar stylo al correo en funcion de unas directrices
  this._addHTMLStyle = function (mensaje) {

    //Logger.log('MENSAJE A FORMATEAR: '+mensaje)
    //cambiar cada salto de linea \n por parrafos <p>
    var mensaje_HTML = "", cadena = mensaje;
    while (cadena && cadena.length != 0) {

      var res = /^([^\n]+)\n/.exec(cadena);
      mensaje_HTML = mensaje_HTML + "<p>" + res[1] + "</p>";
      cadena = cadena.substr(res[0].length);
    }
    mensaje = mensaje_HTML;

    //cambiar <r> por resaltes de estilo <span style="color:rgb(61,133,198);font-weight:700;">
    mensaje = mensaje.replace(/<r>/g, '<span style="color:rgb(61,133,198);font-weight:700;">');
    mensaje = mensaje.replace(/<\/r>/g, '</span>');

    //cambiar <b> por resaltes de estilo <span style="font-weight:700;">
    mensaje = mensaje.replace(/<b>/g, '<span style="font-weight:700;">');
    mensaje = mensaje.replace(/<\/b>/g, '</span>');

    //añadir stilo a las listas <u>

    //añadir imagen this._postal_blob
    mensaje = '<p><img src="cid:' + this._postal_blob.getName() + '" width="602" height="72" style="border:none"></p>'
      + mensaje;

    //estilo general font-size:12pt;color:rgb(7,55,99);font-family:Arial;
    return '<div style="font-size:12pt;color:rgb(7,55,99);font-family:Arial;">' + mensaje + '</div>';

  }
  /*********************
     Feedback
  **********************/
  this.mandarfeedback = function (mails_destinatarios, formulario) {
    var text = this._mensajes.getMensajeEmailFeedbackTXT(formulario.getPublishedUrl());

    var options =
    {
      htmlBody: this._addHTMLStyle(this._mensajes.getMensajeEmailFeedback(formulario.getPublishedUrl())),
      inlineImages:
      {
        postal: this._postal_blob
      }
    };
    this._enviarMailDesdeBuzonPersonal(mails_destinatarios.join(','), 'Cuestionario Alumnos', options, text);
  }

  this.mandarCheckOut = function (mails_destinatarios, formulario) {
    var text = this._mensajes.getMensajeEmailCheckoutTXT(formulario.getPublishedUrl());

    var options =
    {
      htmlBody: this._addHTMLStyle(this._mensajes.getMensajeEmailCheckout(formulario.getPublishedUrl())),
      inlineImages:
      {
        postal: this._postal_blob
      }
    };
    this._enviarMailDesdeBuzonPersonal(mails_destinatarios.join(','), this._mensajes.getAsuntoEmail() + ' | Check-out', options, text);
  }

  this.mandarEvaluacionWorkshop = function (mails_destinatarios, evaluados, formulario, workshop) {

    var text = this._mensajes.getMensajeEmailWorkshopTXT(evaluados, formulario.getPublishedUrl(), workshop);
    var options =
    {
      htmlBody: this._addHTMLStyle(this._mensajes.getMensajeEmailWorkshop(evaluados, formulario.getPublishedUrl(), workshop)),
      inlineImages:
      {
        postal: this._postal_blob
      }
    }

    this._enviarMailDesdeBuzonPersonal(mails_destinatarios.join(','), this._mensajes.getAsuntoEmail() + ' | Evaluaciones Workshop de ' + workshop, options, text);

  }
}


function pruebaComunicaciones()
{
  var comunicaciones= new Comunicaciones();
  comunicaciones.mandarRevalida({getNombre:function(){return 'borja';},getEmail:function(){return 'borja.duran.contractor@bbva.com';}},new Date(),[{capability:'Google Script',link:'http://www.google.es'}])

}