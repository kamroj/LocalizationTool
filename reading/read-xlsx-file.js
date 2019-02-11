const app = require('electron').remote;
const path = require('path');
const dialog = app.dialog;
const excel = require('exceljs');
const div = require('../helpers/div-functions');
const helper = require('../helpers/helper');
const words = require('../additional-functionality/count-words');

let workbook = new excel.Workbook();

document.getElementById('readXLSX').onclick = () => {
    initialize();
    
    dialog.showOpenDialog(fileNames => {
        if (fileNames === undefined) {
            div.showMessage("ERROR: File wasn't chosen!", true);
        } else {            
            readFile(fileNames[0]);
        }
    });
};

function readFile(filepath) 
{
    if (path.extname(filepath) !== '.xlsx') 
    {
        div.showMessage('ERROR: Wrong extension of the file!', true);
        return;
    }

    try 
    {
        workbook.xlsx.readFile(filepath).then(() => {
            div.showMessage(`.xlsx file has been loaded: ${filepath}`);
            enableButtons();
            helper.checkDuplicates(workbook);
            words.countWords(workbook);
        });
    } 
    catch (e) 
    {
        div.showMessage('ERROR: Unable to read .xlsx file!', true);
    }
}

function enableButtons() 
{
    div.blockButton(`countWords`, false);
    div.blockButton(`actorDialogs`, false);
    div.blockButton(`actorDialogsCreate`, false);
    div.blockButton('createJson', false);
    div.blockButton(`createTXT`, false);
    div.blockButton(`replace`, false);
}

function initialize()
{
    words.dispose();
}

module.exports = {
    getXLSXData: () => {
        return workbook;
    }
};
