const XLSX = require('xlsx');

exports.leerExcel= function(ruta){

    function Birthday(mes, dia, nombre, area="", correo){
        this.mes = mes;
        this.dia = dia;
        this.nombre = nombre
        this.area = area;
        this.correo = correo;
    }
    const workbook = XLSX.readFile(ruta);
    const workbookSheets = workbook.SheetNames;
    const regexNumber = /(\d+)/g;//expresión regular ue extrae los numeros de una cadena de texto(solo números, no letras)
    const regexString =/[a-zA-Z\s]/gi; //expresión regular que extrae los meses de una cadena de texto 
    let datosCumpleanos = [];
    const sheet = workbookSheets[0];
    const dataExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], {raw: false});

    
    
    for(const itemFila of dataExcel){
        if(itemFila['Cumpleaños']!=undefined && itemFila['Correo']!=undefined){//si  no hay un campo vacio en el Excel
            let dia = itemFila['Cumpleaños'].match(regexNumber);
            let parseDia = parseInt(dia[0]);
            let date = itemFila['Cumpleaños'].match(regexString);
            let dateSplited = date.join('').trim().toUpperCase();
            let today = new Date();
            let mesEsp = today.toLocaleString('es-ES', { month: 'long' })
            let mesIng = today.toLocaleString('en-EN', { month: 'long' })
            let diaActual = today.getDate();
            let nombre = itemFila["Nombre"].trim();
            let area = itemFila["Área"].trim();
            let correo = itemFila["Correo"].trim();

            if( diaActual == parseDia && dateSplited.substr(0,3) == mesEsp.substr(0,3).toUpperCase() || dateSplited.substr(0,2) == mesIng.substr(0,3).toUpperCase()  ){
                nuevaFecha = new Birthday(dateSplited,dia, nombre, area, correo);
                datosCumpleanos.push(nuevaFecha);
                
            }
            // nuevaFecha = new Birthday(dateSplited,dia, nombre);
            // datosCumpleanos.push(nuevaFecha);
                
            
            // console.log(dateSplited.substr(0,3) + 'Hoy:' + diaEsp.substr(0,3));
            //Obtenemos los meses de los cumpleaños

            
            
            
            //console.log("Desde funcion:" + nuevaFecha)
            
            // console.log('Dias:' + itemFila['Cumpleaños'].match(regexNumber));
            // console.log('Mes:' + dateSplited.trim() );
        }
        

        
    }
    return datosCumpleanos;
    
    
}