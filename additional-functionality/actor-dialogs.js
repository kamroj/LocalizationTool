const xlsxData = require('../reading/read-xlsx-file');
const div = require('../helpers/div-functions');
const words = require('../additional-functionality/count-words');
const $ = require('jquery');
const ws = require('../helpers/worksheet-structure');

const workbookHandler = require('../handlers/workbook-handler');
const columnsData = require('../data/column-structure')

const eventBus = require('../handlers/event-bus');
const eventType = require('../data/event-type');

let selectedActors = [];

let workbookGeneral;
let worksheetGeneral;

document.getElementById('actorDialogs').onclick = () => {
    workbookGeneral = xlsxData.getXLSXData();    
    let actorsFound = prepareActors();

    //clear div
    div.clearDiv(`#description`);
    let englishWrods = words.englishWordMap;    

    for (let index = 0; index < actorsFound.length; index++) 
    {
        actorName = actorsFound[index];

        createNewActorCheckbox(actorName, englishWrods.get(actorName));        
    }   
};

document.getElementById('actorDialogsCreate').onclick = () => {    
    createArrayWithSelectedActors();

    let workbook = workbookHandler.createNewWoorkBook(columnsData.actorDialogs);
    let worksheetNew = workbook.getWorksheet(1);

    fillWorkBook(worksheetNew, worksheetGeneral); 

    div.clearDiv(`#description`);

    eventBus.publish(eventType.SaveExcel, workbook);     
};

function createNewActorCheckbox(actorName, words) 
{
    $('<input />', {
        type: 'checkbox',            
        value: actorName,
    }).appendTo("#description");

    $('<label />', {
        text: `${actorName} :: words: ${words}`        
    }).appendTo("#description").append("<br />");
}

function createArrayWithSelectedActors()
{
    $('input[type=checkbox]').each(function () {
        if (this.checked) {
            //console.log($(this).val()); 
            selectedActors.push($(this).val());
        }
    });
}

function prepareActors()
{
    worksheetGeneral = workbookGeneral.getWorksheet(1);
    let tempActors = [];

    worksheetGeneral.eachRow(row =>{
        if (row.number !== 1)
        {
            if (row.getCell(ws.actorColumn).value !== null)
            {
                tempActors.push(row.getCell(ws.actorColumn).value);
            }
        }
    })

    //remove duplicates
    let actorsFound = [...(new Set(tempActors))];
    console.log(`Actors found -> ${actorsFound.length}`);

    return actorsFound;
}

function fillWorkBook(worksheetNew, worksheetGeneral)
{
    const newKeyColumn = 1;
    const newActorColumn = 2;
    const newEnglishColumn = 3;

    lastRow = 2;

    selectedActors.forEach(actor => {
        
        worksheetNew.getCell(lastRow, 1).value = actor;
        worksheetNew.getCell(lastRow, 1).font = {
            color: { argb: '8A0808' },
            size: 12
        }

        lastRow += 1;
        
        worksheetGeneral.eachRow(row =>{

            if (row.getCell(ws.actorColumn).value === actor)
            {
                worksheetNew.getCell(lastRow, newKeyColumn).value = row.getCell(ws.keyColumn).value;
                worksheetNew.getCell(lastRow, newActorColumn).value = row.getCell(ws.actorColumn).value;
                worksheetNew.getCell(lastRow, newEnglishColumn).value = row.getCell(ws.englishColumn).value;

                lastRow += 1;
            }
        })
    });
}
