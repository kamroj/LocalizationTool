const Excel = require('exceljs');
const assert = require('assert');
const words = require('../additional-functionality/count-words');

let filePath = __dirname + `\\test.xlsx`;
let workbook = new Excel.Workbook();

describe('ExcelTest', function() {
  describe('#CountWords')
    it('should equal', function() {
        workbook.xlsx.readFile(filePath).then(() => {            
            let worksheet = workbook.getWorksheet(1);
            // words.countWords(workbook);
            
            // let sumOfWords = 0;
            // words.polishWordsMap.forEach((value, key) =>{
            //   sumOfWords += words.polishWordsMap.get(key);
            // })
            //console.log(sumOfWords);

            assert.equal(worksheet.getCell(1, 1).value, worksheet.getCell(1, 2).value);
        });
    });
});
