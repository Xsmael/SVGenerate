const jsdom = require("jsdom");
const { JSDOM } = jsdom;
var fs = require('fs-extra');
node_xj = require("xls-to-json");

var log = require('noogger');
const { exec } = require('child_process');
var dom;
var intervalHandler;
const interval= 500
var iteration=0;
/**INPUT DATA */
const SVG_TEMPLATE= "voucher-mini-code.svg";
const XLS_FILE=     "vouchers_portail_roll120.xls";
const JSON_OUTPUT=  "out.json";
const EXPORT_DPI=  300;
const ITEMS_PER_PAGE=16;
fs.readFile(SVG_TEMPLATE, function (err, template) {
    if(err)  log.error(err);
    else {
        dom = new JSDOM(template);
        node_xj({
            input: XLS_FILE,  // input xls
            output: JSON_OUTPUT, // output json
          }, function(err, data) {
            if(err) {
              log.error(err);
            } else {

              log.notice(data);
              intervalHandler= setInterval(function () {
                  makeOneBadge(data);
              },interval);
            }
        });   
    } 
});

function makeOneBadge(data) {
    if(iteration == data.length) {
        clearInterval(intervalHandler);
        return;
    }

    for(let i=1; i<= ITEMS_PER_PAGE; i++) {
      var code= data[iteration++].code; 
      let selector= "#code-"+i+">tspan";
      dom.window.document.querySelector(selector).textContent=code.toUpperCase().centerJustify(14,' ');
    }

    let idx= "page-"+(iteration / ITEMS_PER_PAGE);
    if(!code) return;      
                  //generate SVG files
                  
                  var fileName= idx+'.svg';
                  log.notice(fileName);
                  let filePath = 'out/svg/'+fileName;
                  let fileContent=dom.window.document.querySelector('body').innerHTML;
                  fs.writeFile(filePath, fileContent,{encoding:'utf-8'},(err)=>{
                    if(err) log.error("Failure: "+err);
                    else log.notice("success");
                  });

                  //conversion to PNG using inkscape
                  // inkscape out/svg/page-0.svg --export-dpi=200 --export-png= out/png/page-0.png
                  let cmd= 'inkscape out/svg/'+fileName+' --export-dpi='+EXPORT_DPI+' --export-png='+'out/png/'+fileName.replace('.svg','.png');
                  log.debug(cmd);
                  exec(cmd, (err, stdout, stderr) => {
                      if (err)  log.error(`stdout: ${stderr}`);
                      else   log.debug(`stderr: ${stdout}`);
                    });
}


String.prototype.centerJustify = function( length, char ) {
    var i=0;
	var str= this;
	var toggle= true;
    while ( i + this.length < length ) {
      i++;
	  if(toggle)
	  	str = str+ char;
	  else
	  	str = char+str;
	  toggle = !toggle;
    }
    return str;
}
String.prototype.replaceAll = function(search, replacement) {
    var target = this;
    return target.replace(new RegExp(search, 'g'), replacement);
};