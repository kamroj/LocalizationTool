const excel = require('exceljs');
const app = require('electron').remote;
const dialog = app.dialog;

const xlsxData = require('./merge-xlsx-files');
const div = require('../helpers/div-functions');
const helper = require('../helpers/helper');
const ws = require('../helpers/worksheet-structure');

const EventEmitter = require('events');
class MyEmitter extends EventEmitter {}
const myEmitter = new MyEmitter();

let masterExcel;
let slaveExcel;
let masterWorksheet;
let slaveWorksheet;
let lastRow;
let masterMap = new Map();

let duplicatesWhileMerging = 0;
let mergingConflicts = 0;

// const keyColumn = 1;
// const unitColumn = 2;
// const actorColumn = 3;
// const genderColumn = 4;
// const polishColumn = 5;
// const englishColumn = 8;
// const frenchColumn = 9;
// const polishNewColumn = 11;
// const englishNewColumn = 12;
// const frenchNewColumn = 13;


document.getElementById('mergeLocalization').onclick = () => {
    masterExcel = xlsxData.workbookPrime;
    slaveExcel = xlsxData.workbookToMerge;   

    let masterExcelDuplicates = helper.checkDuplicates(masterExcel);
    let slaveExcelDuplicates = helper.checkDuplicates(slaveExcel);

      
    if (masterExcelDuplicates || slaveExcelDuplicates)
    {
        div.showMessage(`First delete duplicates!`, true);
        return;
    }
    else
    {
        keyChecker();
    }
};

myEmitter.on('showSaveDialog', () => {
    try 
    {
        dialog.showSaveDialog(path => {
            if (path === undefined) 
            {
                return;
            }            
            masterExcel.xlsx.writeFile(path + `.xlsx`).then(() => {
                div.showMessage(`File .xlsx has been saved: ${path}`);                
            });
        });
        masterMap.clear();
        printMergingStatistics();
    } 
    catch (error) 
    {
        div.showMessage('ERROR: Unable to save save .xlsx file!', true);
    }   
});

function keyChecker ()
{
    masterWorksheet = masterExcel.getWorksheet(1);
    slaveWorksheet = slaveExcel.getWorksheet(1);
    lastRow = getLastRowNumber(masterWorksheet);    
    
    masterWorksheet.eachRow(masterRow =>{
        //sprawdzic tez duplikaty
        masterMap.set(masterRow.getCell(ws.keyColumn).value, masterRow.number);
    })    

    slaveWorksheet.eachRow(slaveRow =>{        
        if (masterMap.has(slaveRow.getCell(ws.keyColumn).value))
        {            
            valueChecker(masterMap.get(slaveRow.getCell(ws.keyColumn).value), slaveRow.number);
        }
        else
        {
            writeNewLine(slaveRow.number);
        }
    })

    //debug
    // masterMap.forEach((value, key) =>{
    //     console.log(`${key} -> ${value}`);
    // })

    console.log(`Merging finished`)
    myEmitter.emit('showSaveDialog');
}

function valueChecker(masterRow, slaveRow)
{    
    duplicatesWhileMerging += 1;    

    //jeżeli polskie wartości w master i excel są różne
    if (masterWorksheet.getCell(masterRow, ws.polishColumn).value !== slaveWorksheet.getCell(slaveRow, ws.polishColumn).value)
    {
        mergingConflicts += 1;
        
        console.log(`Row ${masterRow}: Found conflict in polish cells -> 
        base: ${masterWorksheet.getCell(masterRow, ws.polishColumn).value} 
        pending: ${slaveWorksheet.getCell(slaveRow, ws.polishColumn).value}`);
        masterWorksheet.getCell(masterRow, ws.polishConflictColumn).value = slaveWorksheet.getCell(slaveRow, ws.polishColumn).value;
        masterWorksheet.getCell(masterRow, ws.englishConflictColumn).value = slaveWorksheet.getCell(slaveRow, ws.englishColumn).value;
        masterWorksheet.getCell(masterRow, ws.frenchConflictColumn).value = slaveWorksheet.getCell(slaveRow, ws.frenchColumn).value;
    }
    else
    {
        //jeżeli angielskie wartości w master i excel są różne
        if (masterWorksheet.getCell(masterRow, ws.englishColumn).value !== slaveWorksheet.getCell(slaveRow, ws.englishColumn).value)
        {
            if (masterWorksheet.getCell(masterRow, ws.englishColumn).value === null)
            {
                mergingConflicts += 1;

                console.log(`Row ${masterRow}: English cell in master was null. New value merged to base -> ${slaveWorksheet.getCell(slaveRow, ws.englishColumn).value} at row ${masterRow}`);
                masterWorksheet.getCell(masterRow, ws.englishColumn).value = slaveWorksheet.getCell(slaveRow, ws.englishColumn).value;            
            }
            else if (slaveWorksheet.getCell(slaveRow, ws.englishColumn).value !== null)
            {
                mergingConflicts += 1;

                console.log(`Row ${masterRow}: Found conflict in english cells -> 
                base: ${masterWorksheet.getCell(masterRow, ws.englishColumn).value} 
                pending: ${slaveWorksheet.getCell(slaveRow, ws.englishColumn).value}`);
                masterWorksheet.getCell(masterRow, ws.englishConflictColumn).value = slaveWorksheet.getCell(slaveRow, ws.englishColumn).value;
            }   
        }

        //jeżeli francuskie wartości w master i excel są różne
        if (masterWorksheet.getCell(masterRow, ws.frenchColumn).value !== slaveWorksheet.getCell(slaveRow, ws.frenchColumn).value)
        {
            if (masterWorksheet.getCell(masterRow, ws.frenchColumn).value === null)
            {
                mergingConflicts += 1;

                console.log(`Row ${masterRow}: French cell in master was null. New value merged to base -> ${slaveWorksheet.getCell(slaveRow, ws.frenchColumn).value}`);
                masterWorksheet.getCell(masterRow, ws.frenchColumn).value = slaveWorksheet.getCell(slaveRow, ws.frenchColumn).value;       
            }
            else if (slaveWorksheet.getCell(slaveRow, ws.frenchColumn).value !== null)
            {
                mergingConflicts += 1;

                console.log(`Row ${masterRow}: Found conflict in french cells -> 
                base: ${masterWorksheet.getCell(masterRow, ws.frenchColumn).value} 
                pending: ${slaveWorksheet.getCell(slaveRow, ws.frenchColumn).value}`);
                masterWorksheet.getCell(masterRow, ws.frenchConflictColumn).value = slaveWorksheet.getCell(slaveRow, ws.frenchColumn).value;
            }
        }
    }
}

function writeNewLine(slaveRow)
{
    lastRow += 1;
    masterWorksheet.getCell(lastRow, ws.keyColumn).value = slaveWorksheet.getCell(slaveRow, ws.keyColumn).value;
    masterWorksheet.getCell(lastRow, ws.unitColumn).value = slaveWorksheet.getCell(slaveRow, ws.unitColumn).value;
    masterWorksheet.getCell(lastRow, ws.actorColumn).value = slaveWorksheet.getCell(slaveRow, ws.actorColumn).value;
    masterWorksheet.getCell(lastRow, ws.genderColumn).value = slaveWorksheet.getCell(slaveRow, ws.genderColumn).value;
    masterWorksheet.getCell(lastRow, ws.polishColumn).value = slaveWorksheet.getCell(slaveRow, ws.polishColumn).value;
    masterWorksheet.getCell(lastRow, ws.englishColumn).value = slaveWorksheet.getCell(slaveRow, ws.englishColumn).value;
    masterWorksheet.getCell(lastRow, ws.frenchColumn).value = slaveWorksheet.getCell(slaveRow, ws.frenchColumn).value;

    masterMap.set(masterWorksheet.getCell(lastRow, ws.keyColumn).value, lastRow);

    //console.log(`Dodanie nowej lini do mastera na końcu arkusza, row: ${lastRow}`);
}

function getLastRowNumber(worksheet) 
{
    let rowCount = 0;

    worksheet.eachRow(row => {
        rowCount = row.number;
    });

    console.log(`Last row number is: ${rowCount}`);
    return rowCount;
}

function printMergingStatistics()
{
    console.warn(`MERGING STATISTICS:
                  Summary of duplicates found in both excels -> ${duplicatesWhileMerging}
                  Merging conflicts found -> ${mergingConflicts}`);
}