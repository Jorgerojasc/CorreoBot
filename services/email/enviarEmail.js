const nodemailer = require("nodemailer");
const pug = require('pug');
const path = require('path');
const excel = require('../../public/scripts/LeerExcel');
const dotenv = require('dotenv');
dotenv.config();

const image_path = path.join(__dirname + '../../../img/','logo-Uaq.png');
const imagen_fondo_path = path.join(__dirname + '../../../img/','imagenFondo.png');
const path_Excel = path.join(__dirname + '../../../config/BD/Dirección - coordinaciones  BD.xlsx');

let dateObj = new Date();
let year = dateObj.getUTCFullYear();


exports.enviarEmail = ()=>{

    try {

        return new Promise((resolve, reject) =>{
            let dataExcel =  excel.leerExcel(path_Excel);

            if(dataExcel.length >0){
                const transporter = nodemailer.createTransport({
                    // host: "smtp.mailtrap.io",//host de mi servidor
                    // port: 2525,
                    // auth: {
                    // user: "8b775d25b149c2",
                    // pass: "1585f951458780"
                    // },
                    // tls:{
                    //     rejectUnauthorized: false
                    // }
                    service:'gmail',
                    host: "smtp.gmail.com",//host de mi servidor
                    port: 587, 
                    secure:false,
                    auth:{
                        user:"",
                        pass:""
                    },
                    tls:{
                        rejectUnauthorized: false
                    }
                });
                for (let index = 0; index < dataExcel.length; index++) {
                    
                    
                    
                    
                    pug.renderFile(__dirname + "../../../views/index.pug", { name: dataExcel[index].nombre, img:'',imgFondo:'', year, area: dataExcel[index].area }, function (err, data) {
                    
                        //Guardamos el html en una variable
                        let template = pug.compile(data);
                    
                        //OBJETO con los cambios que se realizaran en el template del html(en este caso el src de la imagen)
                        const replacements = {
                            img: "cid:logo-uaq",
                            // imgFondo: "cid:img-decorativa"
                        };
                        
                    
                    
                        const mainOptions = {
                            from: "jorgerojas1cc@gmail.com",
                            to: dataExcel[index].correo,
                            subject: "Felicidades!",
                            html: template(replacements),//template del html con PUG
                            attachments:[
                                {
                                    filename: 'logo-Uaq.png',
                                    path: image_path,
                                    cid: 'logo-uaq'
                                },
                                // {
                                //     filename: 'imagenFondo.png',
                                //     path: imagen_fondo_path,
                                //     cid: 'img-decorativa'
                                // }
                            ]
                        }
                        transporter.sendMail(mainOptions, function (err, info) {
                                if (err) {
    
                                    reject(err.message);//Error al enviar el correo
                                } else {
                                    resolve("Correo(s) enviado(s) " + (index+1) + " de " + dataExcel.length);//Se envio correctamente el correo
                                    
                                }
                            });
                        
                        
                    }); 
                    
                    
                }
            }
            else{
                resolve("Última verificación de cumpleaños: " + (new Date()).getHours() + " hrs" +":"+(new Date()).getMinutes()+" minutos :" + (new Date()).getSeconds() +"s");
            }
               
        })
        

    } catch (error) {
        
        mensajeOutputSYS.push(error);
        return mensajeOutputSYS
    }
    

  

}
