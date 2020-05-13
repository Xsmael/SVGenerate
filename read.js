
var Excel = require('exceljs');

//Read a file
var workbook = new Excel.Workbook();
workbook.xlsx.readFile("vouchers_portail_roll100.xlsx").then(function () {
    var worksheet=workbook.getWorksheet(1);
    worksheet.spliceRows(1,7,['code']);
    let rowCount= worksheet.rowCount;
    for(var i= 2; i <= rowCount ; i++) {
        worksheet.getRow(i).getCell(1).value =  worksheet.getRow(i).getCell(1).value.trim();
        console.log("getRow");
    }

    
    workbook.xlsx.writeFile("vouchers_portail_roll103.xls");

});

console.log("ok boo ok");
