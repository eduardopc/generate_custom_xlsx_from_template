var XlsxTemplate = require('xlsx-template');
var fs = require('fs');
var path = require('path');

fs.readFile(path.join(__dirname, 'report.xlsx'), function(err, data) {
  
  var template = new XlsxTemplate(data);

  var result = { values: ["Eduardo", "Jessica" ]}

  template.substitute(1, result);

  var data = template.generate({type: 'uint8array'});

  fs.writeFileSync(path.join(__dirname, 'report2.xlsx'), data);
});