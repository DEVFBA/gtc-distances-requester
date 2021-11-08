const excelToJson = require('convert-excel-to-json');
const axios = require('axios');
var fs = require('fs');
var json2xls = require('json2xls');

const result = excelToJson({
    sourceFile: './files/211028_sacar_direcciones_de_clientes.xlsx'
});

var origen = []
var distancias = []
var origenes = 0;
for(var i=0; i<result.Hoja1.length; i++)
{
    if(result.Hoja1[i].C === "Origen")
    {
        var origen = {
            direccion: result.Hoja1[i].B,
            identificadorCliente: result.Hoja1[i].A
        }
        var destinos = []
        var contador = 0;
        for(var j=i+1; j<result.Hoja1.length; j++)
        {
            if(result.Hoja1[j].C === "Destino")
            {
                destinos[contador] = {
                    tipo: "1",
                    direccion: result.Hoja1[j].B,
                    latitud: "",
                    longitud: "",
                    identificadorCliente: result.Hoja1[j].A
                }
                
            }
            else if(result.Hoja1[j].C === "Origen")
            {
                j=result.Hoja1.length
            }
            contador++
        }
        //console.log(origen)
        //console.log(destinos)
        distancias[origenes] = {
            origen: {
                tipo: "1",
                direccion: origen.direccion,
                latitud: "",
                longitud: "",
                identificadorCliente: origen.identificadorCliente
            },
            destinos: destinos
        }
        origenes ++
    }
}

var data = {
    distancias: distancias
}

//console.log(data)

async function obtenerDistancias(data){
    var url = "http://129.159.99.152/develop-api/api/distances/WT/"

    try{
        let response = await axios({
            method: 'post',
            url: url,
            json: true,
            data: data
        })
        var json = []
        var distancias = response.data.data.distancias
        var contador = 0;
        for(var i=0; i< distancias.length; i++)
        {
            json[contador] = {
                Identificador: distancias[i].origen.identificadorCliente,
                Direccion: distancias[i].origen.direccion[0],
                Tipo: "Origen",
                Latitud: distancias[i].origen.latitud,
                Longitud: distancias[i].origen.longitud,
                Distancia: ""
            }
            contador ++;
            for(var y=0; y<distancias[i].destinos.length; y++)
            {
                var km = distancias[i].destinos[y].distancia
                json[contador] = {
                    Identificador: distancias[i].destinos[y].identificadorCliente,
                    Direccion: distancias[i].destinos[y].direccion,
                    Tipo: "Destino",
                    Latitud: distancias[i].destinos[y].latitud,
                    Longitud: distancias[i].destinos[y].longitud,
                    Distancia: km.replace(" km", "")
                }
                contador++
            }
        }
        var xls = json2xls(json);
        fs.writeFileSync('./files/Distancias_Clientes.xlsx', xls, 'binary');
    } catch(err){
        console.log(err)
        return {
            mensaje: "Error realizar tu solicitud",
            error: err
        }
    }
}

obtenerDistancias(data)
