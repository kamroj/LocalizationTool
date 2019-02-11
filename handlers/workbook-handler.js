const excel = require('exceljs');
const div = require('../helpers/div-functions');
const ws = require('../helpers/worksheet-structure');


function createNewWoorkBook(worksheetData) 
{
    div.showMessage('Creating new woorkbook!');

    let workbook = new excel.Workbook();
    let worksheet = workbook.addWorksheet('MySheet');    

    workbook.creator = 'KR';      
    worksheet.columns = worksheetData;   

    prettyWorkBook(worksheet);

    return workbook;
}

function prettyWorkBook(worksheet) 
{
    worksheet.properties.outlineLevelCol = 2;
    worksheet.properties.defaultRowHeight = 25; 

    worksheet.getRow(ws.keyColumn).eachCell(cell => {
        cell.font = {
            bold: true
        };
        cell.alignment = {
            vertical: 'middle', 
            horizontal: 'center'
        }
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


module.exports = {
    createNewWoorkBook : (worksheetData) => {
        return createNewWoorkBook(worksheetData);
    }
}