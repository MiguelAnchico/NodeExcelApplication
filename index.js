const fs = require('fs');

// Requiring the module
const reader = require('xlsx')

let data = []
let wrongRows = [];
let newJson = [];

function reformatterExcel() {
    const sheets = file.SheetNames;
    const file = reader.readFile('C:/Users/Cascos Power/Downloads/CRITERIOS CLASIFICACION ITEMS Actual.xlsx')

    for(let i = 0; i < sheets.length; i++){
        const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]])
        let previusCode;
        let previusRow;
        
    
        temp.forEach((res) => {
    
            if(res['CODIGO PLAN'] === 'SL'){
                if(res['REFERENCIA PPAL'] != previusCode) {
                    wrongRows.push({
                        previusRow: previusRow,
                        actualRow: res
                    })
                }
            }
    
            if(res['CODIGO PLAN'] === 'TL') {
                data.push(res)
                newRow = {
                    'REFERENCIA PPAL': res['REFERENCIA PPAL'],
                    'CODIGO PLAN': 'CV',
                    'CODIGO MAYOR': 'UNDEFINED',
                    'CODIGO MAYOR_1': 'Sin Asignar'
                }
                data.push(newRow)
            } else {
                data.push(res)
            }
    
            previusCode = res['REFERENCIA PPAL'];
            previusRow = res;
        })
    
    }
    
    data.map((row) => { 
    
        if(row['CODIGO PLAN'] == 'SL'){
            let jsonInsert = true;
            wrongRows.map((compareRow) => {
                if(compareRow.actualRow['REFERENCIA PPAL'] === row['REFERENCIA PPAL']) {
                    wrongRows.map((correctRow) => {
                        if(correctRow.actualRow['REFERENCIA PPAL'] === compareRow.previusRow['REFERENCIA PPAL']) {
                            newJson.push(correctRow.actualRow)
                            jsonInsert = false;
                        };
                    });
                }
            });
    
            if(jsonInsert) {
                newJson.push(row);
            }
        } else {
            newJson.push(row);
        }
    })
    
    let finalArray = []
    let previusIndex;
    
    for(i = 0; i < newJson.length; i++) {
        if(newJson[i]['CODIGO PLAN'] == 'RF') previusIndex = i;
        if(newJson[i]['CODIGO PLAN'] != 'SL' || newJson[previusIndex]['REFERENCIA PPAL'] == newJson[i]['REFERENCIA PPAL']) finalArray.push(newJson[i]);
    }
    
    
    console.log(wrongRows);
    const json = JSON.stringify(finalArray)

    createJsonFile(json, 'newFormat')
}

function createExcel() {
    const file = reader.readFile('C:/Users/POWER_BIKER_DELL/Downloads/Excel Ultimos Productos.xlsx')
    const sheets = file.SheetNames;
    let data = [];
    let linea = [];
    let json = [];

    for(let i = 0; i < sheets.length; i++){
        const item = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]])

        item.forEach((res) => {
            if(!(data.some((row) => row.Item === res.Item))) data.push(res);
        })

        data.map((res) => {
            if(!(linea.some((row) => row === res.REFERENCIA))) linea.push(res.REFERENCIA);
        })
    }

    json = formatterDataJson(data, 'maestro');
    createJsonFile( JSON.stringify(json), 'newObjectWithAll')
}

function formatterDataJson(Json, type){
    let formatterJson = [];
    Json.map((row) => {
        let LN = row.LINEA;
        let MC = row.MARCAS;
        let RF = row.REFERENCIA;
        let SL = row['TIPO DE PRODUCTO'];
        let TL = row.TALLA;
        let CV = row['COLOR VISOR'];

        if(type != 'maestro') {
        switch(LN) {
            case 'ACCES':
                LN = '03';
                break;
            case 'TEXTIL':
                LN = '02';
                break;
            case 'REPUE':
                LN = '05';
                break;
            case 'MALET':
                LN = '03';
                break;
            case 'CASCO':
                LN = '01';
                break;
        }

        switch(MC) {
            case 'SHAFT':
                MC = '002';
                break;
            case 'TOMCAT':
                MC = '048';
                break;
            case 'TECH':
                MC = '021';
                break;
            case 'AGV':
                MC = '012';
                break;
            case 'PINLOCK':
                MC = '049';
                break;
            case 'BULLET':
                MC = '020';
                break;
            case 'ICH':
                MC = '001';
                break;
            case 'HRO':
                MC = '007';
                break;
            case 'LS2':
                MC = '004';
                break;
            case 'SHAFT_PRO':
                MC = '005';
                break;
            case 'X_ONE':
                MC = '015';
                break;
        }

        switch(RF) {
            case 211:
                RF = '0008';
                break;
            case '307_KP':
                RF = '0171';
                break;
            case '501_SP':
                RF = '0002';
                break;
            case '30_LT':
                RF = '0172';
                break;
            case '3110_DOT':
                RF = '0001';
                break;
            case 'SH-211':
                RF = '0008';
                break;
            case 'HRO-3400DV':
                RF = '0046';
                break;
            case '560_XLITE':
                RF = '0175';
                break;
            case 170:
                RF = '0176';
                break;

            case 'NITRO':
                RF = '0023';
                break;
            case 'SH-571':
                RF = '0019';
                break;
            case '320_STREAM_PINLOCK':
                RF = '0031';
                break;
            case 'K-1_MONO':
                RF = '0117';
                break;
            case 'K-3_SV_TOP':
                RF = '0129';
                break;
            case 516:
                RF = '0045';
                break;
            case 'SHPRO-610DV':
                RF = '0038';
                break;
            case 'SH-227TRIAL':
                RF = '0009';
                break;
            case '503_SP/3120_DV':
                RF = '0003';
                break;
            case 101:
                RF = '0005';
                break;
            case 'SH-551':
                RF = '0017';
                break;
            case 'LS2-FF322/FF351/FF352/FF358/FF384/FF396 ':
                RF = '0032';
                break;
            case '545_HUNTER':
                RF = '0178';
                break;
            case 'SH-502':
                RF = '0011';
                break;
            case 'SH-582SP':
                RF = '0022';
                break;
            case 'HRO-511':
                RF = '0049';
                break;
            case '3110_S_DC':
                RF = '0179';
                break;
            case 'SH-410_X-LITE':
                RF = '0010';
                break;
            case 'X-316RW':
                RF = '0133';
                break;
            case 'SH-3700DVPEAK':
                RF = '0180';
                break;
            case 'SH-581EVO':
                RF = '0021';
                break;
            case 'X-3000GT':
                RF = '0135';
                break;
            case 520:
                RF = '0012';
                break;
            case 'HRO-MX03':
                RF = '0181';
                break;
            case 'SHPRO-MX370DV':
                RF = '0182';
                break;
            case 'SH-560':
                RF = '0018';
                break;
            case 'SH-542GT':
                RF = '0015';
                break;
            case 'AGV-K-1_TOP/K5/K3_SV                    ':
                RF = '0115';
                break;
            case 'HRO-MX330DV':
                RF = '0183';
                break;
            case 'SHPRO-4000DV':
                RF = '0039';
                break;
            case '33_LT/48_LT':
                RF = '0173';
                break;
            case 'SHPRO-600DV':
                RF = '0037';
                break;
            case 'SH-526SP':
                RF = '0014';
                break;
            case 'SH-230':
                RF = '0184';
                break;
            case 'SHPRO-620CARBON':
                RF = '0185';
                break;
            case 'HRO-508DOT':
                RF = '0048';
                break;
            case '34_LT':
                RF = '0186';
                break;
            case '40_LT':
                RF = '0174';
                break;
            case 310:
                RF = '0188';
                break;
            case 'SHAFT-                                  ':
                RF = '0187';
                break;
            case 'SHAFT-545                               ':
                RF = '0189';
                break;
            case '503_SP':
                RF = '0003';
                break;
            case 'K-3_SV_MONO':
                RF = '0115';
                break;
            case '3120_DV':
                RF = '0004';
                break;
            case '3110_S':
                RF = '0179';
                break;
            case '50_LT':
                RF = '0190';
                break;
            case 'LS2-FF320/FF353                         ':
                RF = '0031';
                break;
            case '600_RW':
                RF = '0145';
                break;
            case '62_XTREME':
                RF = '0191';
                break;
            case '50_RW':
                RF = '0192';
                break;
            case 'SHPRO-235DV':
                RF = '0036';
                break;
            case 'SHPRO-4100DV':
                RF = '0193';
                break;
            case 'SH-562':
                RF = '0034';
                break;
            case 'K-1_TOP':
                switch(row.GRAFICOS) {
                    case 'SOLELUNA_2015_(002)':
                        RF = '0122';
                        break;
                    case 'ROSSI_MUGELLO_2016_(007)':
                        RF = '0114';
                        break;
                    case 'SOLELUNA_2017_(009)':
                        RF = '0123';
                        break;
                    case 'DREAMTIME_(005)':
                        RF = '0116';
                        break;
                    case 'ELEMENTS_(018)                          ':
                        RF = '0121';
                        break;
                }
                
                break;
            case 'SH-MX33RAPTOR':
                RF = '0194';
                break;
            case 'HRO-3480DV':
                RF = '0047';
                break;
            case 'SH-522':
                RF = '0013';
                break;
            case '3110_DC':
                RF = '0001';
                break;
            case '35RW                                    ':
                RF = '0195';
                break;
            case '31V                                     ':
                RF = '0006';
                break;
            case 'HRO-518DV':
                RF = '0050';
                break;
            case 'SH-580DV':
                RF = '0020';
                break;
            case 'SHPRO-612DV                             ':
                RF = '0196';
                break;
            case 'SH-3910DV                               ':
                RF = '0197';
                break;
            case 'X-500GT':
                RF = '0134';
                break;
            case 353:
                RF = '0033';
                break;
            case 352:
                RF = '0032';
                break;
            case 'K-1_REPLICA':
                RF = '0120';
                break;
            case 'T10':
                RF = '0198';
                break;
            case 313:
                RF = '0007';
                break;
        }

        switch(CV) {
            case 'N/A':
                CV = '00';
                break;
            case 'SM':
                CV = '01';
                break;
            case 'TR':
                CV = '08';
                break;
            case 'IR.AZ':
                CV = '12';
                break;
            case 'AZ':
                CV = '02';
                break;
            case 'SL':
                CV = '03';
                break;
            case 'IR.AM':
                CV = '05';
                break;
            case 'IR.DO':
                CV = '04';
                break;
            case 'DO':
                CV = '07';
                break;
            case 'GR':
                CV = '00';
                break;
            case 'RJ':
                CV = '00';
                break;
            case 'REVO-AZ':
                CV = '02';
                break;
            case 'REVO-MR':
                CV = '09';
                break;
            case 'N                                       ':
                CV = '00';
                break;
            case 'IR':
                CV = '06';
                break;
            case 'VD':
                CV = '11';
                break;
            case 'MR':
                CV = '09';
                break;
            case 'NJ':
                CV = '10';
                break;
        }

        switch(SL) {
            case 'ALFORJAS':
                SL = '0213';
                break;
            case 'RODILLERA':
                SL = '0214';
                break;
            case 'TORNILLO':
                SL = '0502';
                break;
            case 'RIGIDA':
                SL = '0322';
                break;
            case 'VISOR':
                SL = '0313';
                break;
            case 'TORN_ARAND':
                SL = '0502';
                break;
            case 'ABIERTO':
                SL = '0103';
                break;
            case 'ABATIBLE':
                SL = '0101';
                break;
            case 'CONJU_IMPER':
                SL = '0207';
                break;
            case 'GUANTE':
                SL = 'MANUALMENTE';
                break;
            case 'CHAQ_PROTE':
                SL = '0203';
                break;
            case 'INTEGRAL':
                SL = '0102';
                break;
            case 'SIST_VISOR':
                SL = '0504';
                break;
            case 'PINLOCK':
                SL = '0303';
                break;
            case 'CROSS':
                SL = '0104';
                break;
            case 'MULTIPROP':
                SL = '0105';
                break;
            case 'CHAPA':
                SL = '0501';
                break;
            case 'PANTA_IMPER':
                SL = '0219';
                break;
            case 'CHAQ_IMPER':
                SL = '0226';
                break;
        }

        if(TL === "N/A") TL = 'NA'
            formatterJson.push(
                {
                    'REFERENCIA PPAL': row.Item,
                    'CODIGO PLAN': 'LN',
                    'CODIGO MAYOR': LN
                },
                {
                    'REFERENCIA PPAL': row.Item,
                    'CODIGO PLAN': 'MC',
                    'CODIGO MAYOR': MC
                },
                {
                    'REFERENCIA PPAL': row.Item,
                    'CODIGO PLAN': 'RF',
                    'CODIGO MAYOR': RF
                },
                {
                    'REFERENCIA PPAL': row.Item,
                    'CODIGO PLAN': 'SL',
                    'CODIGO MAYOR': SL
                },
                {
                    'REFERENCIA PPAL': row.Item,
                    'CODIGO PLAN': 'TL',
                    'CODIGO MAYOR': TL
                },
                {
                    'REFERENCIA PPAL': row.Item,
                    'CODIGO PLAN': 'CV',
                    'CODIGO MAYOR': CV
                }
            );
        } else {

            let tipoDeProducto = row['TIPO DE PRODUCTO']+ '';
            let marcas = (row.MARCAS != 'N/A' || row.MARCAS != undefined) ? row.MARCAS + '' : '';
            let graficos = (row.GRAFICOS != 'N/A' || row.GRAFICOS != undefined) ? row.GRAFICOS + '' : '';
            let referencia = (row.REFERENCIA != 'N/A' || row.REFERENCIA != undefined) ? row.REFERENCIA + '' : '';
            let colorPrimario = (new String(row['COLOR PRIMARIO'])  != 'N/A' || row['COLOR PRIMARIO'] != undefined) ? row['COLOR PRIMARIO'] + '' : '';
            let colorSecundario = (new String(row['COLOR SECUNDARIO'])  != 'N/A' || row['COLOR SECUNDARIO'] != undefined) ? row['COLOR SECUNDARIO'] + '' : '';
            let talla = (row.TALLA != 'N/A' || row.TALLA != undefined) ? row.TALLA + '' : '';

            const allDesc = row.LINEA + ' ' + tipoDeProducto.trim() + ' ' + marcas.trim() + ' ' + referencia.trim() + ' ' + graficos.trim() + ' ' + colorPrimario.trim() + ' ' + colorSecundario.trim() + ' ' + talla.trim();
            const descItem = '';
            const descCorta = '';

            formatterJson.push(
                {
                    'CodItem': '',
                    'RefPrincipal': row.Item,
                    'DescItem': descItem,
                    'DescCorta': descCorta,
                    'GrupoImpositivo': '',
                    'TipoInventario': '',
                    'UndInventario': '',
                    'UndOrden': '',
                    'Notas': allDesc,
                }
            );
        }
    });
    
    return formatterJson;

}

function createJsonFile(json, name) {
    fs.writeFile("C:/Users/POWER_BIKER_DELL/Downloads/" + name, json, function(err) {
        if(err) {
            return console.log(err);
        }
        console.log("The file was saved!");
    });
}

createExcel()

/*LN
MC
RF
SL
TL
CV*/