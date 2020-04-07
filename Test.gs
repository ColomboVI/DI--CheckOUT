/**************************************************************************\
 * Copyright (C) 2018 by Synergic Partners                                 *
 *                                                                         *
 * author     : Borja Durán                                                *
 * description:                                                            *
 * - funciones para generar los diferentes tipos de tests                  *
 *                                                                         *
 * TODO                                                                    *
 * ====                                                                    *
 * - .....                                                                 *
 * ----------------------------------------------------------------------- *
 * This program is not free software; you can not : (a) copy or use the    *
 * Software in any manner except as expressly permitted by SynergicPartners*
 * (b) transfer, sell, rent, lease, lend, distribute, or sublicense the    *
 * Software to any third party; (c)  reverse engineer, disassemble, or     *
 * decompile the Software; (d) alter, modify, enhance or prepare any       *
 * derivative work from or of the Software; (e) redistribute it and/or     *
 * modify it without prior, written approval from Synergic Partners.       *
\***************************************************************************/


//@preguntas= [{capability:STRING,npreguntas:INT}]

function Test(array_preguntas)
{
     //Logger.log('array_preguntas: '+JSON.stringify(array_preguntas));
try{
    if (!array_preguntas || array_preguntas.length==0)
      throw 'Se intenta crear un nuevo Test con un conjunto no valido de preguntas.'
    this.getFormulario=function()    {    return this._formulario;    }
    this.getExcelAsociado=function()    {    return this._excel;    }

    this._locale_formularios= LocaleFormularios.getInstancia();


    var numero_preguntas_total = array_preguntas.reduce(function(acc, val) { return acc + val.npreguntas; }, 0);
    var capabilities = array_preguntas.map(function (elem) {return elem.capability;});

    //Logger.log('numero_preguntas_total: '+numero_preguntas_total);
    //Logger.log('capabilities:'+JSON.stringify(capabilities));

    formulario_drive = DriveApp.getFileById('1lPrrz3OEWgo2SX1An3NiKvgV8Za3eLetFlivPZe7jro')
                          .makeCopy('TEST',DriveApp.getRootFolder());

    var ss_base=SpreadsheetApp.openById(this.getExcelBase());

    var form = FormApp.openById(formulario_drive.getId());
    form.setLimitOneResponsePerUser(false).setCollectEmail(false);
    form.setTitle(this.getTitulo())
      .setDescription(this.getCabecera(capabilities.map(function(a){return this.getSheetByName(a).getRange(1,2).getValue();},ss_base),numero_preguntas_total));


    var respuesta=form.createResponse();

      for (var j=0; j<array_preguntas.length;j++)
      {
            var ss=ss_base.getSheetByName(array_preguntas[j].capability);

                  if (array_preguntas[j].capability=='Programación')
                  {

                    form.addSectionHeaderItem().setTitle(this._locale_formularios.getTituloPreguntasCheckinProgramacion());
                    var item= form.addListItem().setTitle(this._locale_formularios.getPregunta1CheckinProgramacion()).setRequired(true);
                    item.setChoices(this._locale_formularios.getRespuestas1CheckinProgramacion().map(function(value){ return this.createChoice(value);},item));
                    item= form.addCheckboxItem().setTitle(this._locale_formularios.getPregunta2CheckinProgramacion());
                    item.setChoices(this._locale_formularios.getRespuestas2CheckinProgramacion().map(function(value){ return this.createChoice(value);},item));
                  }

                form.addPageBreakItem().setTitle(ss.getRange(1,2).getValue());

                //calcular el array de numero de posiciones de pregunrtas: 0.. lastRow-1

                var array_posiciones_preguntas=Array.apply(null, Array(ss.getLastRow()-2)).map(function (x, i) { return i; });//Array.apply(null, {'length': 5}).map(Function.call, Number);

                  for (var k=0; k<array_preguntas[j].npreguntas;k++)
                  {
                        Logger.log(j+'  '+ss.getRange(1,2).getValue());
                          var rr=Math.random()
                          var dest=Math.round((10*rr)%(array_posiciones_preguntas.length-1));

                          var pos=array_posiciones_preguntas.splice(dest,1)[0];

                        //Logger.log(dest+' dest '+pos);
                  var choices = ss.getRange(pos+3,2,1,4).getValues()[0];
                  choices.push(this._locale_formularios.getSinConocimientoRespuesta()); //("Sin conocimiento.");
                  //Logger.log(JSON.stringify(choices));

                      var item=form.addMultipleChoiceItem()
                      .setTitle(ss.getRange(pos+3,1).getValue())
                      .setChoiceValues(choices)
                      .showOtherOption(false).setRequired(true);


                        //Logger.log(' item '+item.getTitle());
                        //Logger.log(' contest '+JSON.stringify(item.getChoices()));
                       // Logger.log(' resp '+ss.getRange(pos+3,1+ss.getRange(pos+3,6).getValue()).getValue());
                      respuesta.withItemResponse(item.createResponse(ss.getRange(pos+3,1+ss.getRange(pos+3,6).getValue()).getValue()));
                  }
      }
      form.addPageBreakItem().setTitle(this._locale_formularios.getAgradecimientoMensaje());
      form.setProgressBar(true);
      respuesta.submit();
      form.setCollectEmail(true).setLimitOneResponsePerUser(true);

      this._excel=SpreadsheetApp.create(formulario_drive.getName()+' (respuestas)');
      //enlazamos con formulario
      form.setDestination(FormApp.DestinationType.SPREADSHEET, this._excel.getId());
      this._formulario=form;
      //Logger.log('generarCopiarTest] para '+this._excel.getName()+' con id '+this._excel.getId()+' es destino del formulario '+formulario_drive.getName());



      return this;
      }
      catch(error)
      {
        if (this._formulario)
          DriveApp.getRootFolder().removeFile(DriveApp.getFileById(this._formulario.getId()));
        if(this._excel)
          DriveApp.getRootFolder().removeFile(DriveApp.getFileById(this._excel.getId()));
        throw error;
      }
}



function CheckoutTest()
{
  this.getExcelBase=function(){ return this._locale_formularios.getExcelPreguntasCheckout();}
  this.getCabecera=function(capabilities,numero_preguntas_total){ return this._locale_formularios.getCabeceraTestCheckout(capabilities,numero_preguntas_total);}
  this.getTitulo=function(){ return this._locale_formularios.getTituloFormularioCheckout();}
  Test.call(this,(new TestConfiguration()).getPreguntasTestCheckOut());

  DriveApp.getFileById(this._formulario.getId()).setName('Check-out');
  DriveApp.getFileById(this._excel.getId()).setName('Check-out (respuestas)');

}



//@workshop= STRING
//@evaluado= STRING
function TestWorkshop(workshop,evaluado)
{

    this.getFormulario=function()    {    return this._formulario;    }
    this.getExcelAsociado=function()    {    return this._excel;    }
    var this_sheet = new ThisSheet();
    this._locale_formularios= LocaleFormularios.getInstancia();

    try{
        var rubricas_conf_ss=SpreadsheetApp.openById(this._locale_formularios.getExcelPreguntasTestWorkshop());
        var configuracion_ss = rubricas_conf_ss.getSheetByName('Configuracion');
        var programas_array = configuracion_ss.getRange(2, 1, configuracion_ss.getLastRow(), 2).getValues();
        var i=0;
        while (i<programas_array.length && programas_array[i][0] && programas_array[i][0]!=this_sheet.getPrograma().getNombre())
            i++;

        var cuestiones_ss=rubricas_conf_ss.getSheetByName(programas_array[i][1])
        var table_cuestiones=cuestiones_ss.getRange(1,1,cuestiones_ss.getLastRow(), 6).getValues();
        if (!table_cuestiones || table_cuestiones.length==0)
          throw 'Se intenta crear un nuevo Test con un conjunto no valido de preguntas.'

        var formulario = DriveApp.getFileById('1lPrrz3OEWgo2SX1An3NiKvgV8Za3eLetFlivPZe7jro').makeCopy('Evaluación de Workshop de '+workshop,DriveApp.getRootFolder());
        var form = FormApp.openById(formulario.getId());

        form.setLimitOneResponsePerUser(false).setCollectEmail(false);
        form.setTitle(this._locale_formularios.getTituloTestWorkshop(workshop))
          .setDescription(this._locale_formularios.getCabeceraTestWorkshop());

        form.addMultipleChoiceItem()
        .setTitle(this._locale_formularios.getCabeceraEvaluacionTestWorkshop())
        .setChoiceValues([evaluado])
        .showOtherOption(false).setRequired(true);

        for (var j=1; j<table_cuestiones.length;j++)
        {
            if (table_cuestiones[j][0])
                form.addSectionHeaderItem().setTitle(table_cuestiones[j][0]); //subject

              form.addScaleItem()
              .setTitle(table_cuestiones[j][1]) // pregunta
              .setBounds(1, 5)//puntuacion de 1 a 5
              .setRequired(table_cuestiones[j][5]);//obligatoriedad
        }

        form.addSectionHeaderItem().setTitle(this._locale_formularios.getAgradecimientoMensaje());
        form.setProgressBar(true);
        form.setLimitOneResponsePerUser(true).setCollectEmail(true);

          this._excel=SpreadsheetApp.create(formulario.getName()+' (respuestas)');
          //enlazamos con formulario
          form.setDestination(FormApp.DestinationType.SPREADSHEET, this._excel.getId());
          this._formulario=form;
          //Logger.log('generarCopiarTest] para '+this._excel.getName()+' con id '+this._excel.getId()+' es destino del formulario '+formulario_drive.getName());
          return this;
      }
      catch(error)
      {
        if (this._formulario)
          DriveApp.getRootFolder().removeFile(DriveApp.getFileById(this._formulario.getId()));
        if(this._excel)
          DriveApp.getRootFolder().removeFile(DriveApp.getFileById(this._excel.getId()));
        throw error;
      }
}



function TestConfiguration() {
//esta clase permite obtener para cada tipo de test, la configuración de que capabilities exponer en el test y cuantas preguntas a partir de unos exceles de refenrecnia
  this._programa=(new ThisSheet()).getPrograma();
  this._pais=(new ThisSheet()).getPais();
  if (typeof TestConfiguration.instancia === 'object') {
        return TestConfiguration.instancia;
    }


    this.getPreguntasTestCheckOut=function()
    {
      if (this._array_preguntas_checkout)
        return this._array_preguntas_checkout;

      var programa_pais = this._programa.getNombre()

        // MACROS_1: Se tiene en cuenta el pais de esta hoja de evaluacion para buscar en la configuracion de tests por programa+pais en casos especificos.
        if (this._pais.getAbreviatura()=="TK")
          programa_pais = programa_pais + "-" + this._pais.getAbreviatura();

        var sheet_test=SpreadsheetApp.openById('1yLKUDqkr4ImuH07MWK34y7ElNztoQoQE1-ZxE2mBw9I');
        var sheet_programas=sheet_test.getSheetByName('Programas Check out');
        var table_programas= new Tabla(sheet_programas,1,1,sheet_programas.getLastRow(), sheet_programas.getLastColumn(), 1);
         this._array_preguntas_checkout=table_programas.getFilaComoObjetoValores(table_programas.getNumFilaColumnaIndexValue(programa_pais))
          .valores
          .filter(function(item){return item.valor!=0;})
          .map(function(item){return {'capability':item.item,'npreguntas':item.valor}});
        return this._array_preguntas_checkout;

    }
 }