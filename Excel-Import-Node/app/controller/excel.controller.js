const multer = require('multer');
const path = require('path');
const excelToJson = require('convert-excel-to-json');
const MongoClient = require('mongodb').MongoClient;

//storage name and path
const storage = multer.diskStorage({
    destination: './public/uploads/',
    filename: function (req, file, cb) {
        cb(null, file.fieldname + '-' + Date.now() + path.extname(file.originalname));
    }
});

//init upload
const upload = multer({
    storage: storage,
    limits: { fileSize: 50000 },
    fileFilter: function (req, file, cb) {
        checkFiletype(file, cb);
    }
}).single('myExcel');

//Check file mime type
function checkFiletype(file, cb) {
    const filetypes = /xlsx|xls/;
    // Check ext
    const extname = filetypes.test(path.extname(file.originalname).toLowerCase());
    console.log(extname);

    if (extname) {
        return cb(null, true);
    } else {
        cb('Error: Excel Only!');
    }
}


//Home Page
const index = (req, res) => res.render('index');

//Handle Excel import
const importExcel = (req, res) => {
    upload(req, res, (err) => {
        if (err) {
            res.render('index', {
                msg: err
            });
        } else {
            if (req.file == undefined) {
                res.render('index', {
                    msg: 'Error: No File Selected!'
                });
            } else {
                console.log(req.file);
                importExcelData2MongoDB(req.file.path)
                res.send('Excel Import successfully');
            }
        }
    });
}

function importExcelData2MongoDB(filePath) {
    // -> Read Excel File to Json Data
    const excelData = excelToJson({
        sourceFile: filePath,
        sheets: [{
            // Excel Sheet Name
            name: 'Users',

            // Header Row -> be skipped and will not be present at our result object.
            header: {
                rows: 1
            },

            // Mapping columns to keys
            columnToKey: {
                B: 'name',
                C: 'address',
                D: 'age',
                E: 'phone',
                F: 'email',
            }
        }]
    });

    // Log Excel Data to Console
    console.log(excelData);
    excelSave(excelData);
}

function excelSave(excelData) {
    var url = "mongodb://localhost:27017";
    MongoClient.connect(url, { useNewUrlParser: true, useUnifiedTopology: true }, function (err, client) {
        if (err) throw err;
        var db = client.db('excel-import');
        db.collection("users").insertMany(excelData.Users, (err, res) => {
            if (err) throw err;
            console.log("Number of documents inserted: " + res.insertedCount);
            client.close();
        });

    });
}

module.exports = { index, importExcel };