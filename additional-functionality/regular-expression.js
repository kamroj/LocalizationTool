const app = require('electron').remote;
const dialog = app.dialog;
const xlsxData = require('../reading/read-xlsx-file');
const ws = require('../helpers/worksheet-structure');
const div = require('../helpers/div-functions');

let workbook;
let replaces = 0;
    
let polishCell;
let englishCell;
let frenchCell;

document.getElementById('replace').onclick = () => {    
    workbook = xlsxData.getXLSXData();
    let worksheet = workbook.getWorksheet(1);
    replace(worksheet);  
};

function replace(worksheet)
{
    worksheet.eachRow(row =>{
        replaceForNewLineHandler(row);
        replaceForTabHanlder(row);
    })

    console.log(`All replacements: ${replaces}`);
    showSaveDialog();
}

function replaceForNewLineHandler(row)
{
    polishCell = row.getCell(ws.polishColumn).value;
    englishCell = row.getCell(ws.englishColumn).value;
    frenchCell = row.getCell(ws.frenchColumn).value;

    if (polishCell !== null && polishCell.includes(`\\n`))
    {
        replaces += 1;
        row.getCell(ws.polishColumn).value = replaceAll(polishCell, `\\\\n`, `\n`);
        console.log(`POLISH -> Row ${row.number}: replaced \\n for new lines :: key : ${row.getCell(ws.keyColumn).value}`)
    }
    if (englishCell !== null && englishCell.includes(`\\n`))
    {
        replaces += 1;
        row.getCell(ws.englishColumn).value =  replaceAll(englishCell, `\\\\n`, `\n`);
        console.log(`ENGLISH -> Row ${row.number}: replaced \\n for new lines :: key : ${row.getCell(ws.keyColumn).value}`)
    }
    if (frenchCell !== null && frenchCell.includes(`\\n`))
    {
        replaces += 1;
        row.getCell(ws.frenchColumn).value =  replaceAll(frenchCell, `\\\\n`, `\n`);
        console.log(`FRENCH -> Row ${row.number}: replaced \\n for new lines :: key : ${row.getCell(ws.keyColumn).value}`)
    }
}

function replaceForTabHanlder(row)
{
    polishCell = row.getCell(ws.polishColumn).value;
    englishCell = row.getCell(ws.englishColumn).value;
    frenchCell = row.getCell(ws.frenchColumn).value;

    if (polishCell !== null && polishCell.includes(`\\t`))
    {
        replaces += 1;
        row.getCell(ws.polishColumn).value =  replaceAll(polishCell, `\\\\t`, `\t`);
        console.log(`POLISH -> Row ${row.number}: replaced \\t for new lines :: key : ${row.getCell(ws.keyColumn).value}`)
    }
    if (englishCell !== null && englishCell.includes(`\\t`))
    {
        replaces += 1;
        row.getCell(ws.englishColumn).value =  replaceAll(englishCell, `\\\\t`, `\t`);
        console.log(`ENGLISH -> Row ${row.number}: replaced \\t for new lines :: key : ${row.getCell(ws.keyColumn).value}`)
    }
    if (frenchCell !== null && frenchCell.includes(`\\t`))
    {
        replaces += 1;
        row.getCell(ws.frenchColumn).value =  replaceAll(frenchCell, `\\\\t`, `\t`);
        console.log(`FRENCH -> Row ${row.number}: replaced \\t for new lines :: key : ${row.getCell(ws.keyColumn).value}`)
    }
}

function replaceAll(str, find, replace) 
{
    return str.replace(new RegExp(find, 'g'), replace);
}

function showSaveDialog()
{
    try 
    {
        dialog.showSaveDialog(path => {
            if (path === undefined) {
                return;
            }
            workbook.xlsx.writeFile(path + `.xlsx`).then(() => {
                div.showMessage(`File .xlsx has been saved: ${path}`);                
            });
        });
    } 
    catch (error) 
    {
        div.showMessage('ERROR: Unable to save save .xlsx file!', true);
    }   
}