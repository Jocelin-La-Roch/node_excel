const express = require('express');
const path = require('path');
const fileUpload = require('express-fileupload');
const xlsx = require('xlsx');
const bodyParser = require('body-parser');
const rootDir = require('./path');
var mysql = require('mysql');

const app = express();

app.set('views', __dirname+'/views');
app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({extended : false}));
app.use(bodyParser.json());

/////////////////////////////////////////////CONNECTION TO DATABASE
var dbConn = mysql.createConnection({
    host: 'localhost',
    user: 'root',
    password: '',
    database: 'excel_to_json'
});
dbConn.connect(function(err) {
    if (err) throw err;
    console.log("Connected!");
});
////////////////////////////////////////////////////////

app.get('/', (req, res, next) => {
    res.sendFile(path.join(rootDir, 'index.html'));
});

app.post('/', (req, res, next) => {
    const workBook = xlsx.readFile(req.body.excelFile);

    let workSheet = {};
    for(const sheetName of workBook.SheetNames){

        workSheet[sheetName] = xlsx.utils.sheet_to_json(workBook.Sheets[sheetName]);
        
    }
    console.log("json:\n", JSON.stringify(workSheet["Feuil1"]), "\n\n");

    dbConn.query('SELECT * FROM users', function (error, results, fields) {
        if (error) throw error;
        return res.send({ error: false, data: results, message: 'users list.' });
    });
    
    res.send(JSON.stringify(workSheet["Feuil1"]));
});

app.listen(4000);

