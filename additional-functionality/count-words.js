const ws = require('../helpers/worksheet-structure');

const columnsStructure = require('../data/column-structure');
const skipWords = require('../data/skip-words-structure');  

const workbookHandler = require('../handlers/workbook-handler');
const eventBus = require('../handlers/event-bus');
const eventType = require('../data/event-type');

let polishWordsMap = new Map();
let englishWordsMap = new Map();
let workbook;

document.getElementById('countWords').onclick = () => {
    initialize();
};

function initialize() 
{
    workbook = workbookHandler.createNewWoorkBook(columnsStructure.countWords);
    let worksheet = workbook.getWorksheet(1);

    let sortedPolishWordMap = sortMapByValue(polishWordsMap);
    let sortedEnglishWordMap = sortMapByValue(englishWordsMap);

    fillWorkBook(worksheet, sortedPolishWordMap, sortedEnglishWordMap);
}


function calculateWords(workbook) 
{      
    let worksheet = workbook.getWorksheet(1);    

    console.log('calcuate words');

    worksheet.eachRow(row => {

        if (row.number === 1)
            return;

        if (row.getCell(ws.actorColumn).value !== null) 
        {            
            let numberOfEnglishWords = countWords(row.getCell(ws.englishColumn).value, skipWords.english);
            let numberOfPolishWords = countWords(row.getCell(ws.polishColumn).value, skipWords.polish);

            if (numberOfEnglishWords !== 0) 
            {
                addCountToEnglishMap(row.getCell(ws.actorColumn), numberOfEnglishWords);
            } 
            else if (numberOfPolishWords !== 0) 
            {
                addCountToPolishMap(row.getCell(ws.actorColumn), numberOfPolishWords);
            } 
            else 
            {
                console.error(`Couldn't find text at row ${row.number}`);
            }
        } 
        else 
        { 
            console.warn(`actor was null at row : ${row.number}`) 
        }
    })
    
    displayWordsStatistics(polishWordsMap, englishWordsMap);    
}

function fillWorkBook(worksheet, sortedPolishWordMap, sortedEnglishWordMap) 
{
    const newKeyColumn = 1;
    const newPolishColumn = 2;
    const newEnglishColumn = 3;

    let lastRow = 2;  
    
    sortedEnglishWordMap.forEach((value, key) => {
        let cellActor = worksheet.getCell(lastRow, newKeyColumn);
        let cellEnglish = worksheet.getCell(lastRow, newEnglishColumn);

        cellActor.value = key;
        cellEnglish.value = value;        

        lastRow = lastRow + 1;
    })

    sortedPolishWordMap.forEach((value, key) =>{
        let keyFound = false;

        worksheet.eachRow(row => {

            if(row.number === 1)
                return;
            
            if(row.getCell(newKeyColumn).value === key)
            {                
                row.getCell(newPolishColumn).value = value;
                keyFound = true;
            }           
        })

        if(keyFound === false)
        {
            worksheet.getCell(lastRow, newKeyColumn).value = key;
            worksheet.getCell(lastRow, newPolishColumn).value = value;            
            
            lastRow = lastRow + 1;
        }
    })

    worksheet.eachRow(row =>{
        if(row.number === 1)
            return;

        let sumCell = row.getCell(4);
        sumCell.value = row.getCell(newPolishColumn).value + row.getCell(newEnglishColumn).value;
    })
    
    eventBus.publish(eventType.SaveExcel, workbook)    
}

function addCountToPolishMap(actorCell, numberOfWords) 
{
    if (polishWordsMap.has(actorCell.value)) 
    {        
        polishWordsMap.set(actorCell.value, polishWordsMap.get(actorCell.value) + numberOfWords);        
    } 
    else 
    {
        polishWordsMap.set(actorCell.value, numberOfWords);
    }

    numberOfWords = 0;
}

function addCountToEnglishMap(actorCell, numberOfWords) 
{
    if (englishWordsMap.has(actorCell.value)) 
    {
        englishWordsMap.set(actorCell.value, englishWordsMap.get(actorCell.value) + numberOfWords);
    } 
    else 
    {
        englishWordsMap.set(actorCell.value, numberOfWords);
    }

    numberOfWords = 0;
}

function countWords(cellValue, skipwords) 
{
    if (cellValue === null)
        return 0;

    let regex = /\b\S+\b/g; //dzieli na słowa
    let words = cellValue.match(regex);

    //zwracam 0 inaczej będzie typ NaN - not a number
    if(words === null)
        return 0;

    if (skipwords === null)
        return words.length;

    let w2 = words.filter(function(item)
    {
        return !skipwords.includes(item);
    });

    return w2.length;
}

function displayWordsStatistics(polishWordMap, englishWordMap)
{
    let sortedPolishWordMap = sortMapByValue(polishWordMap);
    let sortedEnglishWordMap = sortMapByValue(englishWordMap);

    //Debug
    console.log(`\n POLSKIE:`)
    sortedPolishWordMap.forEach((value, key) => {        
        console.log(`POLSKIE: klucz: ${key} : ${value}`);
    })

    console.log(`\n ANGIELKSIE:`)
    sortedEnglishWordMap.forEach((value, key) => {        
        console.log(`ANGIELSKIE: klucz: ${key} : ${value}`);
    })
}

function sortMapByValue(map) 
{
    return tempMap = new Map([...map.entries()].sort((a, b) => b[1] - a[1]));
}

function dispose()
{
    polishWordsMap.clear();
    englishWordsMap.clear();
}

module.exports = {
    countWords: (workbook) => {
        calculateWords(workbook);
    },

    polishWordsMap: polishWordsMap,
    englishWordMap: englishWordsMap,

    dispose: () =>{
        dispose();
    }
}
