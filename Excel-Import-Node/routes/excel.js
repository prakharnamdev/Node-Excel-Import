const express = require('express');
const router = express.Router();
const {index,importExcel} = require('../app/controller/excel.controller');

router.get('/',index );
router.post('/upload',importExcel);

module.exports=router