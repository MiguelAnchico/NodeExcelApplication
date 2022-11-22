const fs = require('fs');

// Requiring the module
const reader = require('xlsx')
// Reading our test file
const file = reader.readFile('C:/Users/Cascos Power/Downloads/CRITERIOS CLASIFICACION ITEMS Actual.xlsx')
let data = []
const sheets = file.SheetNames;

let wrongRows = [];
let newJson = [];

for(let i = 0; i < sheets.length; i++){
    const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]])
    let previusCode;
    let previusRow;
    

    temp.forEach((res) => {

        if(res['CODIGO PLAN'] == 'SL'){
            if(res['REFERENCIA PPAL'] != previusCode) {
                wrongRows.push({
                    previusRow: previusRow,
                    actualRow: res
                })
            }
        }

        if(res['CODIGO PLAN'] == 'TL') {
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
            if(compareRow.actualRow['REFERENCIA PPAL'] == row['REFERENCIA PPAL']) {
                wrongRows.map((correctRow) => {
                    if(correctRow.actualRow['REFERENCIA PPAL'] === compareRow.previusRow['REFERENCIA PPAL']) newJson.push(correctRow.actualRow);
                });
                jsonInsert = false;
                
            }
        });

        if(jsonInsert) {
            newJson.push(row);
        }
    } else {
        newJson.push(row);
    }
})


console.log(newJson);
const json = JSON.stringify(data)

fs.writeFile("C:/Users/Cascos Power/Desktop/newClient", json, function(err) {
    if(err) {
        return console.log(err);
    }
    console.log("The file was saved!");
});