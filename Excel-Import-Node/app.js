const express = require('express');
const ejs = require('ejs');
const app = express();
const excel = require('./routes/excel');

app.set('view engine','ejs');

app.use(express.static('./public'));

app.use('/',excel);

app.listen(5000,()=>{
    console.log('port listening to 5000 ');
    
})