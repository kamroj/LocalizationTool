const app = require('electron').remote;
const path = require('path');
const dialog = app.dialog;
const excel = require('exceljs');
const div = require('../helpers/div-functions');
const json = require('../reading/read-json-file');
const ws = require('../helpers/worksheet-structure');

let workbookPrime = new excel.Workbook();
let workbookToMerge = new excel.Workbook();

document.getElementById('masterXlsx').onclick = () => {
    dialog.showOpenDialog((fileNames) => {

        if (fileNames === undefined) 
        {
            div.showMessage('ERROR: File wasn\'t chosen!', true);
        }
        else 
        {
            readFile(fileNames[0], workbookPrime);
        }
    });
};

document.getElementById('slaveXlsx').onclick = () => {
    dialog.showOpenDialog((fileNames) => {
        if (fileNames === undefined) 
        {
            div.showMessage('ERROR: File wasn\'t chosen!', true);
            return;
        }
        else 
        {
            readFile(fileNames[0], workbookToMerge);
        }
    });
};

document.getElementById('mergeXlsx').onclick = () => {
    mergeFiles();
    try {
        dialog.showSaveDialog((path) => {
            if (path === undefined) 
                return;

            workbookPrime.xlsx.writeFile(path + '.xlsx').then(() => {
                div.showMessage(`File has been saved: ${path}`);
            })
        });
    }
    catch (e) 
    {
        div.showMessage('ERROR: Unable to save save .xlsx file!', true);
    }
};

document.getElementById('mergeActors').onclick = () => {
    let jsonData = json.getJsonData();

    if (jsonData == undefined) 
    {
        div.showMessage('ERROR: Actor JSON is not loaded! Please load actor JSON first.', true);
        return;
    } 
    else if (jsonData.actorLocalizationStrings == null) 
    {
        div.showMessage('ERROR: Wrong JSON file! Please load actor JSON file then try again.', true);
        return;
    }

    mergeActorsToXlsx(jsonData);
    try 
    {
        dialog.showSaveDialog((path) => {
            if (path === undefined) { return; };
            workbookPrime.xlsx.writeFile(path + '.xlsx').then(() => {
                div.showMessage(`File .xlsx has been saved: ${path}`);
            })
        });
    }
    catch (e) 
    {
        div.showMessage('ERROR: Unable to save save .xlsx file!', true);
    }
}

function readFile(filepath, workbook) 
{
    if (path.extname(filepath) !== '.xlsx') 
    {
        div.showMessage('ERROR: Wrong extension of the file!', true);
        return;
    }
    try 
    {
        workbook.xlsx.readFile(filepath).then(() => {
            if (workbook == workbookPrime) 
            {
                document.getElementById('masterXlsx').setAttribute("type", "mergeButtonLoaded");
                div.showMessage(`Prime .xlsx file is loaded: ${filepath}`);
                div.blockButton('mergeActors', false);
            }
            if (workbook == workbookToMerge) 
            {
                document.getElementById('slaveXlsx').setAttribute("type", "mergeButtonLoaded")
                div.showMessage(`Slave .xlsx file to merge is loaded: ${filepath}`);
            }
            if (workbookPrime.getWorksheet() != undefined && workbookToMerge.getWorksheet() != undefined) 
            {
                div.showMessage('Merge button is unlocked!');
                div.blockButton('mergeXlsx', false);
                div.blockButton('mergeLocalization', false)
            }
        });
    }
    catch (e) 
    {
        div.showMessage('ERROR: Unable to read XLSX file!', true);
    }
}

function mergeFiles() 
{
    let worksheetPrime = workbookPrime.getWorksheet(1);
    let worksheetToMerge = workbookToMerge.getWorksheet(1);

    worksheetToMerge.eachRow(rowToMerge => {
        if (rowToMerge.number != 1) 
        {
            worksheetPrime.eachRow(rowPrime => {                
                if (rowPrime.getCell(ws.keyColumn).value === rowToMerge.getCell(ws.keyColumn).value) 
                {
                    rowPrime.getCell(ws.polishColumn).value == rowToMerge.getCell(ws.polishColumn).value ?
                        rowPrime.getCell(ws.statusColumn).value = " " :
                        rowPrime.getCell(ws.statusColumn).value = "changed";

                    if (rowToMerge.getCell(ws.actorColumn).value !== null && rowPrime.getCell(ws.actorColumn).value !== rowToMerge.getCell(ws.actorColumn).value)
                    {
                        rowPrime.getCell(ws.actorColumn).value === null ? 
                        console.log(`Row ${rowPrime.number}: add new actor ${rowToMerge.getCell(ws.actorColumn).value}`) :
                        console.warn(`Row ${rowPrime.number}: Found actor conflict!
                            was: ${rowPrime.getCell(ws.actorColumn).value}
                            changed to: ${rowToMerge.getCell(ws.actorColumn).value}`);
                            
                        rowPrime.getCell(ws.actorColumn).value = rowToMerge.getCell(ws.actorColumn).value;
                        // if (rowPrime.getCell(ws.actorColumn).value === null)
                        // {
                        //     rowPrime.getCell(ws.actorColumn).value = rowToMerge.getCell(ws.actorColumn).value;
                        //     console.log(`Row ${rowPrime.number}: add new actor ${rowToMerge.getCell(ws.actorColumn).value}`);
                        // }
                        // else
                        // {
                        //     rowPrime.getCell(ws.englishConflictColumn).value = rowToMerge.getCell(ws.actorColumn).value;
                        //     console.warn(`Row ${rowPrime.number}: Found actor conflict!
                        //         base: ${rowPrime.getCell(ws.actorColumn).value}
                        //         pending: ${rowToMerge.getCell(ws.actorColumn).value}`);
                        // }
                    }
      
                    rowPrime.getCell(ws.englishColumn).value = rowToMerge.getCell(ws.englishColumn).value;
                    rowPrime.getCell(ws.frenchColumn).value = rowToMerge.getCell(ws.frenchColumn).value;
                }
            })
        }
    })
    div.showMessage('Merge has been complited!');
}

function mergeActorsToXlsx(json) 
{
    let worksheetPrime = workbookPrime.getWorksheet(1);

    worksheetPrime.eachRow(row => {
        json.actorLocalizationStrings.forEach(item => {

            if (row.getCell(ws.keyColumn).value === item.key) 
            {
                if (row.getCell(ws.actorColumn).value !== item.actorName && row.getCell(ws.actorColumn).value !== null)
                {
                    console.warn(`Row ${row.number}: change actor name from ${row.getCell(3).value} -> ${item.actorName}`);
                    row.getCell(ws.actorColumn).value = item.actorName;
                }
                else if (row.getCell(ws.actorColumn).value === null)
                {
                    console.log(`Row ${row.number}: added new actor name -> ${item.actorName}`);
                    row.getCell(ws.actorColumn).value = item.actorName;
                }
                
                if (row.getCell(ws.genderColumn).value !== item.actorGender && row.getCell(ws.genderColumn).value !== null)
                {
                    console.warn(`Row ${row.number}: change actor gender from ${row.getCell(ws.genderColumn).value} -> ${item.actorGender}`);
                    row.getCell(ws.genderColumn).value = item.actorGender;
                }
                else if (row.getCell(ws.genderColumn).value === null)
                {
                    console.log(`Row ${row.number}: added new actor gender -> ${item.actorGender}`);
                    row.getCell(ws.genderColumn).value = item.actorGender;
                }                
            }
        })
    })
}

module.exports = {
    workbookPrime: workbookPrime,     
    workbookToMerge : workbookToMerge    
}