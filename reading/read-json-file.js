const app = require('electron').remote;
const path = require('path');
const dialog = app.dialog;
const fs = require('fs');
const div = require('../helpers/div-functions');

let data;

document.getElementById('readJson').onclick = () => {
    dialog.showOpenDialog(fileNames => {        
        if (fileNames === undefined) 
        {
            div.showMessage("ERROR: File wasn't chosen!", true);
        } 
        else 
        {
            readFile(fileNames[0]);
        }
    });
};

function readFile(filepath) 
{
    fs.readFile(filepath, 'utf-8', (err, jsonData) => {
        if (err || path.extname(filepath) !== '.json') 
        {
            div.showMessage('ERROR: Wrong extension of the file!', true);
            return;
        }
        try 
        {
            data = JSON.parse(jsonData);
            console.log(data);
            div.blockButton('createXLSX', false);
            div.showMessage(`Json has been loaded from: ${filepath}`);
        } 
        catch (e) {
            div.showMessage('ERROR: Unable to parse .json file!', true);
        }
    });
}

module.exports = {
    getJsonData: () => {
        return data;
    }
};
