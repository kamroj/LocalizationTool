const excel = require('exceljs');
const app = require('electron').remote;
const dialog = app.dialog;
const json = require('../reading/read-json-file');
const div = require('../helpers/div-functions');
const prompt = require('electron-prompt');

const ws = require('../helpers/worksheet-structure');
const columnsStructure = require('../data/column-structure'); 
const workbookHandler = require('../handlers/workbook-handler');

const EventEmitter = require('events');
class MyEmitter extends EventEmitter {}
const myEmitter = new MyEmitter();

let jsonData;
let columnLanguage = 5;
let thirdLanguageName = 'FRENCH';
let workbook;

document.getElementById('createXLSX').onclick = () => {
    workbook = initialize();
};

myEmitter.on('showSaveDialog', () => {
    try 
    {
        dialog.showSaveDialog(path => {
            if (path === undefined)
                return;
            
                workbook.xlsx.writeFile(path + '.xlsx').then(() => {
                div.showMessage(`File .xlsx has been saved: ${path}`);
            });
        });
    } 
    catch (e) 
    {
        div.showMessage('ERROR: Unable to save save .xlsx file!', true);
    }
});

function initialize() 
{
    div.showMessage('Creating new woorkbook!');

    columnsStructure.localization[8].header = thirdLanguageName;
    workbook = workbookHandler.createNewWoorkBook(columnsStructure.localization);
    let worksheet = workbook.getWorksheet(1);

    // workbook.creator = 'KR';
    // worksheet.properties.outlineLevelCol = 2;
    // worksheet.properties.defaultRowHeight = 15;

    // worksheet.columns = [
    //     { header: 'ID', key: '1', width: 30 },
    //     { header: 'UNIT', key: '2', width: 20 },
    //     { header: 'ACTOR', key: '3', width: 10 },
    //     { header: 'GENDER', key: '4', width: 10 },
    //     { header: 'POLISH', key: '5', width: 50 },
    //     { header: 'STATUS', key: '6', width: 10 },
    //     { header: 'COMMENTARY', key: '7', width: 30 },
    //     { header: 'ENGLISH', key: '8', width: 50 },
    //     { header: thirdLanguageName, key: '9', width: 50 },
    //     { header: 'PACK', key: '10', hidden: true }
    // ];

    fillWorkBook(worksheet);
    //prettyWorkBook(worksheet);
    checkDuplicates(worksheet);
    return workbook;
}

function fillWorkBook(worksheet) 
{
    jsonData = json.getJsonData();
    let row = 2;
    
    jsonData.localizationPacks.forEach(package => {
        package.localizationUnits.forEach(element => {
            element.localizationStrings.forEach(item => {
                let cellId = worksheet.getCell(row, 1);
                let cellUnit = worksheet.getCell(row, 2);
                let cellLanguage = worksheet.getCell(row, columnLanguage);
                let cellStatus = worksheet.getCell(row, 6);
                let cellPack = worksheet.getCell(row, 10);

                cellId.value = item.key;
                cellUnit.value = element.name;
                cellLanguage.value = item.value;
                cellStatus.value = 'new';
                cellPack.value = package.name;

                row += 1;
                });
            });
        });   
}

function prettyWorkBook(worksheet) 
{
    worksheet.getRow(ws.keyColumn).eachCell(cell => {
        cell.font = {
            bold: true
        };
    });
    worksheet.getColumn(ws.polishColumn).eachCell(cell => {
        cell.alignment = {
            wrapText: true
        };
    });
    worksheet.getColumn(ws.englishColumn).eachCell(cell => {
        cell.alignment = {
            wrapText: true
        };
    });
    worksheet.getColumn(ws.frenchColumn).eachCell(cell => {
        cell.alignment = {
            wrapText: true
        };
    });
}

function checkDuplicates(worksheet) 
{
    const duplicateString = `duplicated`;
    let duplicatesArray = [];
    let duplicatesFound = false;

    worksheet.getColumn(ws.keyColumn).eachCell(cell => {
        if (duplicatesArray.includes(cell.value)) 
        {
            div.showMessage(`Duplicated key: ${cell.value}`, true);
            worksheet.getCell(cell.row, ws.commentColumn).value = duplicateString;
            duplicatesFound = true;
        }
        duplicatesArray.push(cell.value);
    });

    if (duplicatesFound) 
    {
        dialog.showMessageBox(
            {
                type: 'question',
                buttons: ['OK', 'Cancel'],
                message: 'Duplicates detected, would you like to delete them?'
            },
            buttons => {
                
                if (buttons == 0) 
                {
                    let rowCount = 0;
                    worksheet.getColumn(ws.commentColumn).eachCell(() => {
                        rowCount++;
                    });                    

                    for (var row = 0; row <= rowCount; row++) 
                    {
                        if (worksheet.getCell(row, ws.commentColumn).value == duplicateString) 
                        {
                            //jeden oznacza ile ma usunąć wierszy
                            worksheet.spliceRows(row, 1);
                            row -= 1;
                        }
                    }
                    myEmitter.emit('showSaveDialog');
                } 
                else 
                {
                    myEmitter.emit('showSaveDialog');
                }
            }
        );
    } else {
        myEmitter.emit('showSaveDialog');
    }
}

//should i do it in another class?
const $ = require('jquery');
let isLanguageSet = false;
var createXLSXPL = document.getElementById('createXLSXPL');
var createXLSXENG = document.getElementById('createXLSXENG');
var createXLSXFR = document.getElementById('createXLSXFR');

createXLSXPL.addEventListener('click', () => {
    columnLanguage = 5;
    div.showMessage('.xlsx language is set to: POLISH');
});

createXLSXENG.addEventListener('click', () => {
    columnLanguage = 8;
    div.showMessage('.xlsx language is set to: ENGLISH');
});

createXLSXFR.addEventListener('click', () => {
    isLanguageSet
        ? div.showMessage(`.xlsx language is set to: ${thirdLanguageName}`)
        : myEmitter.emit('showPromptLanguage');
    columnLanguage = 9;
});

myEmitter.once('showPromptLanguage', () => {
    prompt({
        height: 150,
        label: 'Please, type language name',
        value: ''
    }).then(value => {
        if (value === null || value === '') {
            div.showMessage(`.xlsx language is set to default: ${thirdLanguageName}`);
            return;
        } else {
            thirdLanguageName = value.toLocaleUpperCase();
            div.showMessage(`.xlsx language is set to: ${thirdLanguageName}`);
            isLanguageSet = true;
        }
    });
});

//changing visual of selected buttons
$('.xlsxLanguage').click(function() {
    $(this).prop('type', 'langButtonSelected');
    $('.xlsxLanguage')
        .not(this)
        .prop('type', 'langButton');
});
