/***************************************************************************************
LISTADO DE FUNCIONES de la clase Mensajes (ESPAÑOL->MensajesES // INGLES->MensajesEN ):
 ------------------------------
 + getMensajeEmailCheckout
 + getMensajeEmailCheckoutTXT
***************************************************************************************/

var Mensajes = (function() {
  var instance;
  function createInstancia() {
    return new ThisSheet().getPais().isLenguajeEnglish() ? new MensajesEN() : new MensajesES();
  }
  return {
    getInstancia: function() {
      if (!instance) {
        instance = createInstancia();
      }
      return instance;
    }
  };
})();

var meses = [
  'enero',
  'febrero',
  'marzo',
  'abril',
  'mayo',
  'junio',
  'julio',
  'agosto',
  'septiembre',
  'octubre',
  'noviembre',
  'diciembre'
];
var dias = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
var meses_EN = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December'
];
var dias_EN = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

function MensajesES() {
  this._programa = new ThisSheet().getPrograma();
  this._getColetilla = function() {
    return '<b>¡Aprovecha</b> <r>tu oportunidad!</r>\n';
  };
  this._getColetillaFeedback = function() {
    return '<b>¡tu</b> <r>opinión es muy importante para nosotros!</r>\n';
  };
  this._eliminarFormato = function(text) {
    return text.replace(/<[^>]+>/g, '');
  };
  this._getFechaLocale = function(fecha) {
    return dias[fecha.getDay()] + ' ' + fecha.getDate() + ' de ' + meses[fecha.getMonth()];
  };
  this.getAsuntoEmail = function() {
    return 'Campus BBVA: Programa ' + this._programa.getNombre();
  };

  this.getMensajeEmailCheckout = function(link_formulario) {
    return (
      'Hola,\nComo alumno del <r>Programa ' +
      this._programa.getNombre() +
      '</r> debes responder al siguiente <a href="' +
      link_formulario +
      '">test de comprobación de conocimiento</a>.\n' +
      this._getColetilla()
    );
  };
  this.getMensajeEmailCheckoutTXT = function(link_formulario) {
    return (
      'Hola,\n\nComo alumno del Programa ' +
      this._programa.getNombre() +
      ' debes responder al siguiente test de comprobación de conocimiento (' +
      link_formulario +
      ').\n'
    );
  };

  this.getMensajeEmailFeedback = function(link_formulario) {
    return (
      'Hola,\nComo alumno del <r>Programa ' +
      this._programa.getNombre() +
      '</r> nos gustaria conocer tu opinión, contestando al siguiente  <a href="' +
      link_formulario +
      '">formulario</a>.\n' +
      this._getColetillaFeedback()
    );
  };

  this.getMensajeEmailFeedbackTXT = function(link_formulario) {
    return (
      'Hola,\n\nComo alumno del Programa ' +
      this._programa.getNombre() +
      ' nos gustaria conocer tu opinión (' +
      link_formulario +
      ').\n'
    );
  };

  this.getMensajeEmailWorkshop = function(evaluados, link_formulario, workshop) {
    return (
      'Hola,\nCompleta el <a target="_blank" href="' +
      link_formulario +
      '">siguiente formulario de valoración</a> del trabajo del <b>workshop de ' +
      workshop +
      '</b> realizado por: <r>' +
      evaluados +
      '</r>\n'
    );
  };
  this.getMensajeEmailWorkshopTXT = function(evaluados, link_formulario, workshop) {
    return (
      'Hola,\n\nCompleta el siguiente formulario de valoración (' +
      link_formulario +
      ') del trabajo del workshop de ' +
      workshop +
      ' realizado por: ' +
      evaluados
    );
  };
}

function MensajesEN() {
  this._programa = new ThisSheet().getPrograma();
  this._getColetilla = function() {
    return '<b>Sieze</b> <r>the opportunity!</r>\n';
  };
  this._eliminarFormato = function(text) {
    return text.replace(/<[^>]+>/g, '');
  };
  this._getFechaLocale = function(fecha) {
    return (
      dias_EN[fecha.getDay()] + ' ' + meses_EN[fecha.getMonth()] + ' ' + fecha.getDate() + 'th'
    );
  };
  this.getAsuntoEmail = function() {
    return 'BBVA Campus: ' + this._programa.getNombre() + ' Program';
  };

  this.getMensajeEmailCheckout = function(link_formulario) {
    return (
      'Hello.\nAs a student of the <r>' +
      this._programa.getNombre() +
      ' Program</r> you must answer the following <a href="' +
      link_formulario +
      '">knowledge test</a>.\n' +
      this._getColetilla()
    );
  };
  this.getMensajeEmailCheckoutTXT = function(link_formulario) {
    return this._eliminarFormato(this.getMensajeEmailCheckout(link_formulario));
  };

  this.getMensajeEmailWorkshop = function(evaluados, link_formulario, workshop) {
    return (
      'Hello,\nComplete the <a target="_blank" href="' +
      link_formulario +
      '">following form</a> that will you allow to evaluate the work done for the <b>' +
      workshop +
      ' workshop</b> carried out by: <r>' +
      evaluados +
      '</r>\n'
    );
  };
  this.getMensajeEmailWorkshopTXT = function(evaluados, link_formulario, workshop) {
    return (
      'Hello,\n\nComplete the following form (' +
      link_formulario +
      ') that will you allow to evaluate the work done for the ' +
      workshop +
      ' workshop carried out by: ' +
      evaluados
    );
  };
}
