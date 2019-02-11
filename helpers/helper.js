function checkDuplicates(workbook) 
{
    const keyColumn = 1;

    let keyHolder = new Map();
    let worksheet = workbook.getWorksheet(1);
    let foundDuplicates = false; 
    
    worksheet.eachRow(row => {
        if (keyHolder.has(row.getCell(keyColumn).value)) 
        {
            foundDuplicates = true;
            console.log(`Duplicates: ${row.getCell(keyColumn).value} at rows ${keyHolder.get(row.getCell(keyColumn).value)} and ${row.number}`);
        } 
        else 
        {
            keyHolder.set(row.getCell(keyColumn).value, row.number);            
        }
    })
    keyHolder.clear();
    return foundDuplicates;
}

module.exports = {
    checkDuplicates : (workbook) =>{
        return checkDuplicates(workbook)
    }
}