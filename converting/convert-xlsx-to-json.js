const excel = require('exceljs');
const app = require('electron').remote;
const dialog = app.dialog;
const xlsxData = require('../reading/read-xlsx-file');
const c = require('../helpers/LocalizationClasses');
const fs = require('fs');
const div = require('../helpers/div-functions');
const ws = require('../helpers/worksheet-structure');

let workbook;
let data;
let columnLanguage = 5;

document.getElementById('createJson').onclick = () => {
    prepareJson();
    try 
    {
        dialog.showSaveDialog(path => {
            if (path === undefined) {
                return;
            }
            writeFile(path + '.json', data);
        });
    } 
    catch (error) 
    {
        div.showMessage('ERROR: Unable to save save .json file!', true);
    }
};

function writeFile(filepath, content) 
{
    fs.writeFile(filepath, content, err => {
        if (err) 
        {
            div.showMessage('ERROR: Unable to safe file!', true);
        }
        div.showMessage(`JSON has been created: ${filepath}`);
    });
}

function prepareJson() {
    let worksheet;
    workbook = xlsxData.getXLSXData();
    worksheet = workbook.getWorksheet(1);

    var container = new c.LocalizationContainer();
    let pack = new c.LocalizationPack();
    var unit = new c.LocalizationUnit();
    var string;
    var map = new Map();    

    pack.name = worksheet.getRow(2).getCell(ws.packColumn).value;
    unit.name =
        worksheet.getRow(2).getCell(ws.unitColumn).value == null
            ? 'null'
            : worksheet.getRow(2).getCell(ws.unitColumn).value;

    container.containerName = 'TestContainerName';

    map.set(unit.name, unit);

    worksheet.eachRow(row => {
        if (row.number != 1) 
        {
            if (row.getCell(ws.packColumn).value != pack.name && row.getCell(ws.packColumn).value != null) 
            {
                container.localizationPacks.push(pack);
                pack = new c.LocalizationPack();
                pack.name = row.getCell(ws.packColumn).value;
            }

            if (row.getCell(ws.unitColumn).value != unit.name) 
            {
                if (map.has(row.getCell(ws.unitColumn).value)) 
                {                    
                    unit = map.get(row.getCell(ws.unitColumn).value);
                } 
                else 
                {
                    unit = new c.LocalizationUnit();
                    unit.name = row.getCell(ws.unitColumn).value;
                    map.set(unit.name, unit);
                }
            }

            string = new c.LocalizationString();

            row.eachCell((cell, colNumber) => {
                if (colNumber == 1) 
                {
                    string.key = cell.value;
                }                
                if (colNumber == columnLanguage) 
                {
                    string.value = cell.value;
                }
            });

            if (string.key && string.value) 
            {
                validateLinks(string.key, string.value);
                console.log(`Push strings to object: ${unit.name}`);
                unit.localizationStrings.push(string);
            } 
            else 
            {
                console.warn(`Skip empty row: ${row.number}`);
            }
        }
    });   
   
    for (var value of map.values()) 
    {
        pack.localizationUnits.push(value);
        console.log(`${value.name} unit has been added.`);
    }

    container.localizationPacks.push(pack);

    data = JSON.stringify(container, null, 4);
}

function validateLinks(key, value) 
{
    if (key.includes('_DOC_')) 
    {
        let openLinks = (value.match(/<a name/g) || []).length;
        let closingLinks = (value.match(/<\/a>/g) || []).length;

        if (openLinks != closingLinks) 
        {
            div.showMessage(`Incorect links: ${key}`, true);
        }
    }
}

//ugly
const $ = require('jquery');
var createJsonPL = document.getElementById('createJsonPL');
var createJsonENG = document.getElementById('createJsonENG');
var createJsonFR = document.getElementById('createJsonFR');

createJsonPL.addEventListener('click', () => {
    columnLanguage = 5;
    div.showMessage('Json language is set to: POLISH');
});

createJsonENG.addEventListener('click', () => {
    columnLanguage = 8;
    div.showMessage('Json language is set to: ENGLISH');
});

createJsonFR.addEventListener('click', () => {
    columnLanguage = 9;
    div.showMessage('Json language is set to: 3rd language');
});

//changing visual of selected buttons
$('.jsonLanguage').click(function () {
    $(this).prop('type', 'langButtonSelected');
    $('.jsonLanguage')
        .not(this)
        .prop('type', 'langButton');
});
