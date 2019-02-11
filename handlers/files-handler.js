const app = require('electron').remote;
const dialog = app.dialog;

const div = require('../helpers/div-functions');
const eventBus = require('./event-bus');
const eventType = require('../data/event-type')

eventBus.subscribe(eventType.SaveExcel, workbooks => saveExcelFile(workbooks));

function saveExcelFile(file)
{
    try 
    {
        dialog.showSaveDialog(path => {
            if (path === undefined)
                return;

            file.xlsx.writeFile(path + `.xlsx`).then(() => {
                div.showMessage(`File .xlsx has been saved: ${path}`);                
            });
        });
    } 
    catch (error) 
    {
        div.showMessage('ERROR: Unable to save save .xlsx file!', true);
    }   
}
