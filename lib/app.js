//Dependencies
const _data = require('./data');


lib = {}

lib.fillData = (dir, jsonName, xlsxName, worksheetName) => {

    // reading a data from json file
    _data.read(dir, jsonName)
        .then(data => {
            const table = [];
            for (let i = 0; i < data.length; i+=4) {
                const helpArr = [];
                for (let j = i; j < i + 4; j++) {
                    helpArr.push(data[j]);
                }
                table.push(helpArr);
            }
            return table;
        })
        .then(data => {
            // reading and filling xlsx file
            _data.readAndFill(dir, xlsxName, worksheetName, data)
        })
        .catch(err => console.log('ERROR: Problem with reading a json file'));
}

// Export module
module.exports = lib;