const xlsx = require('xlsx');
const workBook = xlsx.readFile("test.xlsx"); 
const workSheet = workBook.Sheets[workBook.SheetNames[0]];

const titles = [];
let argument = {};

const jsonResult = [];
let document = {};

for (let cell in workSheet){
    const cellAsString = cell.toString();

    if(cellAsString[1] !== 'r' && cellAsString !== 'm' && cellAsString[1] == 1){
        if(cellAsString [0] === 'A'){
            argument.title = workSheet[cell].v;
        }
        if(cellAsString [0] === 'B'){
            argument.author = workSheet[cell].v;
        }
        if(cellAsString [0] === 'C'){
            argument.released = workSheet[cell].v;
            titles.push(argument);
            argument = {};
        }
    }
}

for (let cell in workSheet){
    const cellAsString = cell.toString();

    if(cellAsString[1] !== 'r' && cellAsString !== 'm' && cellAsString[1] > 1){
        if(cellAsString [0] === 'A'){
            document[titles[0].title] = workSheet[cell].v;
        }
        if(cellAsString [0] === 'B'){
            document[titles[0].author] = workSheet[cell].v;
        }
        if(cellAsString [0] === 'C'){
            document[titles[0].released] = workSheet[cell].v;
            jsonResult.push(document);
            document = {};
        }
    }
}
console.log(titles);
console.log(jsonResult);