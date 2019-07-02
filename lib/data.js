//Dependencies
const fs = require('fs');
const path = require('path');
const Excel = require('exceljs/modern.nodejs');
const workbook = new Excel.Workbook();

const lib = {};

lib.baseDir = path.join(__dirname, '/../.data')

// Read file 

lib.read = (dir, file) => {
    return new Promise((res, rej) => {
        fs.readFile(lib.baseDir + dir + '/' + file, 'utf-8', (err, data) => {
            if(!err && data) {
                const obj = JSON.parse(data);
                res(obj)
            } else {
                rej(err)
            }
        })
    })
}

lib.readAndFill = (dir, file, worksheetName, data) => {

    // Reading file xlsx without data
    workbook.xlsx.readFile(lib.baseDir + dir + '/' + file)
        .then(() => {
            let worksheet = workbook.getWorksheet(worksheetName);
            if (worksheet) {
                
                let amountOfRows = worksheet.rowCount;
                let count = 0;
                
                for (let i = 0; i < amountOfRows; i++) {
                    let row = worksheet.getRow(i);

                    if (row.getCell(1).value === 'temperatura') {
                        row.getCell(3).value = data[count][0]['THBTemperature'];
                        worksheet.getRow(i + 1).getCell(3).value = data[count][0]['THBHumidity'];
                        worksheet.getRow(i + 2).getCell(3).value = data[count][0]['THBPressure'];

                        for (let j = 0; j < 4; j++) {
                            row.getCell(4 + j).value = data[count][j]['mass_in_g'];
                        }

                        row.getCell(10).value = data[count][3]['THBTemperature'];
                        worksheet.getRow(i + 1).getCell(10).value = data[count][3]['THBHumidity'];
                        worksheet.getRow(i + 2).getCell(10).value = data[count][3]['THBPressure'];
                        count++;
                    }
                }
                    
                for (let i = 2; i < 8; i++) {
                    worksheet.getColumn(i).width = 15
                }
                
                worksheet.getColumn(10).width = 15;

                // Writing the xlsx file with data
                let newFileName = 'filled__' + Date.now() + '.xlsx'
                workbook.xlsx.writeFile(newFileName)
                    .then(() => console.log("A file is saved on: " + path.join(__dirname, '/../') + newFileName))
                    .catch(() => console.log("ERROR: Could not save a file"))
            } else {
                console.log('ERROR: Worksheet not exist')
            }
        })
        .catch(err => console.log('ERROR: Problem with reading a xlsx file'));
    
}

// Export module
module.exports = lib;