const fs = require('fs');
const xlsxData = require('../reading/read-xlsx-file');
const ws = require('../helpers/worksheet-structure');


document.getElementById('createTXT').onclick = () => {
    init();
};

let dataMap = new Map();
let unitMap = new Map();

function init()
{
    let woorkbook = xlsxData.getXLSXData();
    let worksheet = woorkbook.getWorksheet(1);
    prepareMaps(worksheet, ws.englishColumn);
    prepareMaps(worksheet, ws.frenchColumn);

}

function prepareMaps(worksheet, column)
{
    worksheet.eachRow(row => {
        if (row.number === 1)
            return;
        
        let unitName = row.getCell(ws.unitColumn).value;
        let key = row.getCell(ws.keyColumn).value;
        let cellValue = row.getCell(column).value;        
        
        if(cellValue === null)
            return;

        if (unitMap.has(unitName))
        {
            let tempData = unitMap.get(unitName);
            tempData.set(key, cellValue);
        }
        else
        {
            dataMap = new Map();
            dataMap.set(key, cellValue);    
            unitMap.set(unitName, dataMap);        
        }
    })
    
    writeTxtFiles(worksheet.getCell(1, column).value);
}


function writeTxtFiles(folderName) 
{
    unitMap.forEach((data, unitName) => {
        let fileName = unitName;
        let fileData = ``;
        data.forEach((value, key) => {
            fileData += key + `\n` + value + `\n`;
        });

        if (!fs.existsSync(folderName)){
            fs.mkdirSync(folderName);
        }
        fs.writeFile(folderName + `/` + fileName + '.txt', fileData, function (err, fileData) {
            err ? console.log(err) : console.log(`success`);
        });
    });

    unitMap.clear();
    dataMap.clear();
}

