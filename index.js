const nodeZip = require('node-zip');
const constants = require('./constants.json');
const Sheet = require('./sheet');

const sheetsFront = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
+ '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
+ '<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/><workbookPr defaultThemeVersion="124226"/><bookViews><workbookView xWindow="360" yWindow="90" windowWidth="28035" windowHeight="12345" activeTab="0"/></bookViews><sheets>';
const sheetsBack = '</sheets><calcPr calcId="144525"/></workbook>';
const relFront = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
const relBack = '<Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId102" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme.xml"/></Relationships>';
const contentTypeFront = '<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />'
  + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />'
  + '<Default Extension="psmdcp" ContentType="application/vnd.openxmlformats-package.core-properties+xml" />'
  + '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />'
  + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />'
  + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />';
const contentTypeBack = '</Types>';

function generateMultiSheets (configs, xlsx) {
  configs.forEach(function (config, i) {
    var index = i + 1;
    config.name = config.name ? config.name : (`sheet${index}`);

    var sheet = new Sheet(config, xlsx);
    sheet.generate();
  });
}

function generateContentType(configs, xlsx) {
  var workbook = contentTypeFront;
  
  configs.forEach(function (config) {
    workbook += `<Override PartName="/${config.fileName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />`;
  });

  workbook += contentTypeBack;
  xlsx.file('[Content_Types].xml', workbook);
}

function generateRel(configs,xlsx) {
  var workbook = relFront;

  configs.forEach(function (config, i) {
    var index = i + 1;
    workbook += `<Relationship Id="rId${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${config.name}.xml"/>`;
  });

  workbook += relBack;
  xlsx.file('xl/_rels/workbook.xml.rels', workbook);
  xlsx.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
   + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' 
   + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' 
   + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
}

function generateWorkbook (configs,xlsx) {
  var workbook = sheetsFront;

  configs.forEach(function(config, i) {
    var index = i + 1;
    workbook += `<sheet name="${config.name}" sheetId="${index}" r:id="rId${index}"/>`;
  });

  workbook += sheetsBack;
  xlsx.file('xl/workbook.xml', workbook);
}

exports.execute = function(config) {
  var configs = [];
  if (config instanceof Array) {
    configs = config;
  } else {
    configs.push(config);
  }

  var xlsx = nodeZip(constants.xlsTemplate, { base64: true, checkCRC32: false });
  generateMultiSheets(configs, xlsx);
  generateWorkbook(configs, xlsx);
  generateRel(configs, xlsx);
  generateContentType(configs, xlsx); 

  return xlsx.generate({
    base64: false,
    compression: 'DEFLATE'
  });
};

