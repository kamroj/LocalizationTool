const { google } = require("googleapis");
const drive = google.drive("v3");
const key = require("./private-key.json");
const path = require("path");
const fs = require("fs");
const div = require('../helpers/div-functions');
const app = require("electron").remote;
const dialog = app.dialog;

const EventEmitter = require("events");
class MyEmitter extends EventEmitter {}
const myEmitter = new MyEmitter();

let folderId = "1ckox60_yZh5Ss1f9JT6OSsaivMp3NSjw";

var jwtClient = new google.auth.JWT(
    key.client_email,
    null,
    key.private_key,
    ["https://www.googleapis.com/auth/drive"],
    null
);

document.getElementById("google").onclick = () => {
    myEmitter.emit("getAuthorization");
    dialog.showOpenDialog(fileNames => {
        if (fileNames == undefined) 
            return;

        readFile(fileNames[0]);
    });
};

myEmitter.once("getAuthorization", () => {
    jwtClient.authorize(authErr => {        
        if (authErr) 
        {
            div.showMessage(`ERROR: Authorization failed: ${authErr}`, true);
            return;
        } 
        else 
        {
            div.showMessage("Authorization succeed!");
        }
    });
});

function readFile(filepath) {
    fs.readFile(filepath, "utf-8", err => {
        let pathSplit = filepath.split("\\");
        let fileName = pathSplit[pathSplit.length - 1];

        var fileMetadata = {
            name: fileName,
            parents: [folderId]
        };
        var media = {
            mimeType: setGoogleSheetFormat(filepath),
            body: fs.createReadStream(filepath)
        };
        drive.files.create(
            {
                auth: jwtClient,
                resource: fileMetadata,
                media: media,
                fields: "id"
            },
            err => {
                if (err) 
                {
                    div.showMessage(`Error: ${err}`);
                } 
                else 
                {
                    div.showMessage(`File ${fileName} has been uploaded: 
                    https://drive.google.com/drive/u/0/folders/1ckox60_yZh5Ss1f9JT6OSsaivMp3NSjw`);
                }
            }
        );
    });
}

function setGoogleSheetFormat(filepath) 
{
    switch (filepath) 
    {
        case `.xlsx`:
            return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";         
        case `.txt`:
            return "text/plain";        
        case `.json`:
            return "text/plain";        
        case `.docx`:
            return "application/vnd.oasis.opendocument.text";         
        case `.zip`:
            return "application/zip";            
        case `.pdf`:
            return "application/pdf";            
        default:            
            break;
    } 
}
