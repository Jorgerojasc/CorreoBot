
const express = require("express");
const app = express();
const excel = require('./public/scripts/LeerExcel');
const email = require('./services/email/enviarEmail');
const cron = require('node-cron');
const dotenv = require('dotenv');
dotenv.config();
const port = process.env.PORT || 3000;

//cron ejecutara el script en una cierta hora o fecha elegida
cron.schedule('00 00 * * *', () => {// LOS ASTERISCO SIGNIFICAN EN ESTE ORDEN ( MINUTOS, HORAS, DIA DEL MES, MES , DIA DE LA SEMANA) 
    email.enviarEmail()
    .then((result) => {

        console.log(result)
    }).catch((err) => {
        
        console.log(err);
    });
});

    
    
app.post('/send-email',(req, res)=>{
    // setInterval(() => {
    //     email.enviarEmail()
    //     .then((result) => {
    //         res.status(200).jsonp(req.body);;
    //         console.log(result)
    //     }).catch((err) => {
    //         res.status(500).jsonp(req.body);;
    //         console.log(err);
    //     });
    // }, 1000);
   email.enviarEmail()
    .then((result) => {
        res.status(200).jsonp(req.body);;
        console.log(result)
    }).catch((err) => {
        res.status(500).jsonp(req.body);;
        console.log(err);
    });
    
});

app.post("/Excel", (req, res)=>{
    //leerExcel('./config/BD/Dirección BD.xlsx');
    var data = excel.leerExcel('./config/BD/Dirección BD.xlsx');
    //console.log(data[0].dia[0]);
    for (let index = 0; index < data.length; index++) {
        console.log("Dia: " + data[index].dia[0] + " Mes: " + data[index].mes + "Nombre:" + data[index].nombre + "Area: " + data[index].area)
        
    }
    res.status(200).jsonp(req.body);
})
app.listen(port, process.env.HOST, ()=>{
    console.log(`Servidor en -> http://${process.env.HOST}:${process.env.PORT}`);
})
