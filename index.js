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
    
    fs.writeFile("C:/Users/Cascos Power/Desktop/newClient", json, function(err) {
        if(err) {
            return console.log(err);
        }
        console.log("The file was saved!");
    });
}

function createExcel() {
    const file = reader.readFile('C:/Users/Cascos Power/Desktop/Excel Ultimos Productos.xlsx')
    const sheets = file.SheetNames;
    let data = [];
    let linea = [];

    for(let i = 0; i < sheets.length; i++){
        const item = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]])

        item.forEach((res) => {
            if(!(data.some((row) => row.Item === res.Item))) data.push(res);
        })

        data.map((res) => {
            if(!(linea.some((row) => row === res['TIPO DE PRODUCTO']))) linea.push(res['TIPO DE PRODUCTO']);
        })
    }

    console.log(formatterDataJson(data));

}

function formatterDataJson(Json){
    let formatterJson = [];
    Json.map((row) => {
        let LN = row.LINEA;
        let MC = row.MARCAS;
        let RF = row.REFERENCIA;
        let SL = row['TIPO DE PRODUCTO'];
        let TL = row.TALLA;
        let CV = row['COLOR VISOR'];

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
            case 'SHAFT':
                RF = '002';
                break;
            case 'TOMCAT':
                RF = '048';
                break;
            case 'TECH':
                RF = '021';
                break;
            case 'AGV':
                RF = '012';
                break;
            case 'PINLOCK':
                RF = '049';
                break;
            case 'BULLET':
                RF = '020';
                break;
            case 'ICH':
                RF = '001';
                break;
            case 'HRO':
                RF = '007';
                break;
            case 'LS2':
                RF = '004';
                break;
            case 'SHAFT_PRO':
                RF = '005';
                break;
            case 'X_ONE':
                RF = '015';
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
                'CODIGO MAYOR': row.REFERENCIA
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

    });

    return formatterJson;

}

createExcel()

/*LN
MC
RF
SL
TL
CV*/