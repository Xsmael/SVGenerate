var xlsx= require("xlsx");

var file= xlsx.readFile("participants-170718.xls");
  

var out= xlsx.utils.sheet_to_json(xlsx.utils.aoa_to_sheet(file));
console.log(out);
