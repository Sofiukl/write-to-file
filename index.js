const fs = require('fs');
const Path = require('path');
const Excel = require('exceljs');

let workbook = null;
let worksheet = null;


let startRow = 0;
let endRow = 0;

const startCol = 'B';
const endCol = 'Q';
const imageHeight = 22;
const gapBtwImg = 5;


/**
Create the excel file 
**/
fs.readdir(Path.join(__dirname, 'test-cases'), (err, items) => {
	console.log(`No of items : ${items.length}`);

	createExcel(items);

});

/**
Creating excel file

**/
let createExcel = (items) => {

	_createWorkBook();
	_addImages(items);
	_saveExcel();
	
}

let _createWorkBook = () => {
	try {
        workbook = new Excel.Workbook();
        worksheet = workbook.addWorksheet('Test Execution');

        /*worksheet.columns = [
            { header: 'Id', key: 'id', width: 10 },
            { header: 'Name', key: 'name', width: 32 },
            { header: 'D.O.B.', key: 'DOB', width: 10 }
        ];
        worksheet.addRow({id: 1, name: 'SofikulM', DOB: new Date(2018,4,22)});
        worksheet.addRow({id: 2, name: 'Jane Doe', DOB: new Date(2018,4,22)});  */  

    } catch(err) {
        console.log(`Fail to create the Excel: ${err}`);
    }
}

let _addImages = (items) => {

	items.forEach( (i) => {
		console.log(Path.join(__dirname, 'test-cases', i));
		_addImage(Path.join(__dirname, 'test-cases', i), 'png');
	})
}

let _addImage = (fileName, ext) => {
	var img = workbook.addImage({
  		filename: fileName,
  		extension: ext,
	});

	console.log(img);
	
	
	startRow = endRow + gapBtwImg;
	endRow = startRow + imageHeight;

	let imgPosition = `${startCol + startRow}:${endCol + endRow}`;

	worksheet.addImage(img, imgPosition);
}

let _saveExcel = () => {
	
	workbook.xlsx.writeFile("./test-cases/Test Evidence.xlsx").then(function() {
    	console.log("Excel(xlsx) file is written.");
	});
}