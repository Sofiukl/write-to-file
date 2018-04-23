const fs = require('fs');
const Path = require('path');

const Excel = require('exceljs');

let workbook = null;
let worksheet = null;
/**
Creating excel file

**/
let createExcel = () => {

	try {
        workbook = new Excel.Workbook();
        worksheet = workbook.addWorksheet('Test Execution');

        worksheet.columns = [
            { header: 'Id', key: 'id', width: 10 },
            { header: 'Name', key: 'name', width: 32 },
            { header: 'D.O.B.', key: 'DOB', width: 10 }
        ];
        worksheet.addRow({id: 1, name: 'SofikulM', DOB: new Date(2018,4,22)});
        worksheet.addRow({id: 2, name: 'Jane Doe', DOB: new Date(2018,4,22)});    

    } catch(err) {
        console.log('OOOOOOO this is the error: ' + err);
    }
}

let addImage = (fileName, ext) => {
	var img = workbook.addImage({
  		filename: fileName,
  		extension: ext,
	});

	console.log(img);

	// insert an image over part of B2:D6
	worksheet.addImage(img, 'B4:O24');
}

let saveExcel = () => {
	
	workbook.xlsx.writeFile("./test-cases/Test Evidence.xlsx").then(function() {
    		console.log("Excel(xlsx) file is written.");
	});
}


fs.readdir(Path.join(__dirname, 'test-cases'), (err, items) => {
	console.log(`No of items : ${items.length}`);

	createExcel();

	items.forEach( (i) => {
		console.log(Path.join(__dirname, 'test-cases', i));
		addImage(Path.join(__dirname, 'test-cases', i), 'png');
	})

	saveExcel();
});