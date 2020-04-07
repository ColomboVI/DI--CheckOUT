/**************************************************************************\
 * Copyright (C) 2018 by Synergic Partners                                 *
 *                                                                         *
 * author     : Borja Durán                                                *
 * description:                                                            *
 * - clases que agrupan por idioma determinado texto a la hora de generar los test      *
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


var LocaleFormularios = (function () {
    var instance;
 
    function createInstancia() {
        var this_sheet= new ThisSheet();        
        return (this_sheet.getPais().isLenguajeEnglish())?(new LocaleFormulariosEN()):(new LocaleFormulariosES());
    }
 
    return {
        getInstancia: function () {
            if (!instance) {
                instance = createInstancia();
            }
            return instance;
        }
    };
})();


function LocaleFormulariosES ()
{

    this._programa= (new ThisSheet()).getPrograma();
    this.getAgradecimientoMensaje= function()    {        return '¡Gracias por tu participación!';    }
    this.getSinConocimientoRespuesta= function()    {        return 'no lo sé.';    }
    
   
    //CHECK OUT
    this.getExcelPreguntasCheckout= function()    {        return '1yLKUDqkr4ImuH07MWK34y7ElNztoQoQE1-ZxE2mBw9I';    }
    this.getCabeceraTestCheckout= function(capabilities,numero_preguntas_total)    {        return 'El test contiene '+numero_preguntas_total+' preguntas con 4 opciones de respuesta (sólo una es correcta).\r\n'+
            'Evaluaremos tus conocimientos actuales en:\r\n\r\n'+
            capabilities.join('\r\n')+
            '\r\n\r\n'+
            'Por favor, NO busques las respuestas en Internet.\r\n'+
            'El objetivo es tener una visión clara de tus conocimientos actuales.';    }
    this.getTituloFormularioCheckout= function()    {        return 'Programa '+this._programa.getNombre()+' - Proceso de Check-out';    }

               
    //TestWorkshop
    this.getExcelPreguntasTestWorkshop= function()    {        return '1EIT6pTBAJZcu3bFNDeTHPu1IMqei7TqgEnJ5RDNwh68';    }
    this.getCabeceraTestWorkshop= function()    {        return 'Este cuestionario te permitirá evaluar el workshop realizado por tus compañeros. Valora teniendo en cuenta la rúbrica.';    }
    this.getTituloTestWorkshop= function(workshop)    {        return 'Programa '+this._programa.getNombre()+' - Evaluación de Workshop de '+workshop;    }
    this.getCabeceraEvaluacionTestWorkshop=function(){return 'Vas a evaluar a los siguientes candidatos:';}
               
    
}


function LocaleFormulariosEN ()
{
    this._programa= (new ThisSheet()).getPrograma();

    this.getAgradecimientoMensaje= function()    {        return 'Thanks for your participation!';    }
    this.getSinConocimientoRespuesta= function()    {        return 'I don\'t know.';    }

    
    
    //CHECK OUT
    this.getExcelPreguntasCheckout= function()    {     return '1DJP2gvp2_s54XD7WlbNQxbIN0q4kP09GF7opRn9K4fE';    }
    this.getCabeceraTestCheckout= function(capabilities,numero_preguntas_total)    {        return 'The test contains '+numero_preguntas_total+' questions with 4 answer options (only one is correct).\r\n'+
            'We will evaluate your current knowledge in:\r\n\r\n'+
            capabilities.join('\r\n')+
            '\r\n\r\n'+
            'Please, do NOT look for the answers on the Internet.\r\n'+
            'The objective is to have a clear vision of your current knowledge to help you reach the necessary level before the program.\r\n\r\n';    }
    this.getTituloFormularioCheckout= function()    {        return this._programa.getNombre()+' Program - Check-out process';    }    
               
    //TestWorkshop
    this.getExcelPreguntasTestWorkshop= function()    {        return '1ZhcgA0WHcmB8Z9Vk2qlV9fwGt--BBYbFGy5GAymKn1w';    }
    this.getCabeceraTestWorkshop= function()    {        return 'This questionnaire will allow you to evaluate the workshop carried out by your classmates. You\'ll have to value taking into account the rubric.';    }
    this.getTituloTestWorkshop= function(workshop)    {        return this._programa.getNombre()+' Program - '+workshop+' workshop evaluation';    }
    this.getCabeceraEvaluacionTestWorkshop=function(){return 'You are going to evaluate the following candidates:';}
    
}