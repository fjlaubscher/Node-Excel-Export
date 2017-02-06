var sheetFront = '<?xml version="1.0" encoding="utf-8"?><x:worksheet xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
	+ ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
	+ '<x:sheetPr/><x:sheetViews><x:sheetView tabSelected="1" workbookViewId="0" /></x:sheetViews>'
	+ '<x:sheetFormatPr defaultRowHeight="15" />';
var sheetBack = '<x:pageMargins left="0.75" right="0.75" top="0.75" bottom="0.5" header="0.5" footer="0.75" /><x:headerFooter /></x:worksheet>';
    
var fs = require('fs');

// helpers
var helpers = require('./helpers');
var addNumberCell = helpers.addNumberCell;
var addDateCell = helpers.addDateCell;
var addBoolCell = helpers.addBoolCell;
var addStringCell = helpers.addStringCell;
var getColumnLetter = helpers.getColumnLetter;

function Sheet (config, xlsx) {
  this.config = config;
  this.xlsx = xlsx;
}

Sheet.prototype.generate = function() {
  var config = this.config;
  var xlsx = this.xlsx;
  var cols = config.cols;
  var data = config.rows;

  var rows = '';
  var row = '';
  var colsWidth = '';
  var self = this;

  config.fileName = 'xl/worksheets/' + (config.name || 'sheet').replace(/[*?\]\[\/\/]/g, '') + '.xml';

  if (config.stylesXmlFile) {
    var styles = fs.readFileSync(config.stylesXmlFile, 'utf8');
    if (styles) {
      xlsx.file('xl/styles.xml', styles);
    }
  }

	//first row for column caption
  row = '<x:row r="1" spans="1:' + cols.length + '">';
  
  for (var k = 0; k < cols.length; k++) {
    var colStyleIndex = cols[k].captionStyleIndex || 0;
    row += addStringCell(self, getColumnLetter(k + 1) + 1, cols[k].caption, colStyleIndex);
    
    if (cols[k].width) {
      colsWidth += '<col customWidth = "1" width="' + cols[k].width + '" max = "' + (k + 1) + '" min="' + (k + 1) + '"/>';
    }
  }

  row += '</x:row>';
  rows += row;

	//fill in data
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var currRow = i + 2;
    row = '<x:row r="' + currRow + '" spans="1:' + cols.length + '">';
    
    for (var j = 0; j < cols.length; j++) {
      var cellData = r[j];
      var cellType = cols[j].type;
      var styleIndex = null;
      
      if (typeof cols[j].beforeCellWrite === 'function') {
        var e = {
          rowNum: currRow,
          styleIndex: null,
          cellType: cellType
        };

        cellData = cols[j].beforeCellWrite(r, cellData, e);
        styleIndex = e.styleIndex || styleIndex;
        cellType = e.cellType;
      }

      switch (cellType) {
      case 'number':
        row += addNumberCell(getColumnLetter(j + 1) + currRow, cellData, styleIndex);
        break;
      case 'date':
        row += addDateCell(getColumnLetter(j + 1) + currRow, cellData, styleIndex);
        break;
      case 'bool':
        row += addBoolCell(getColumnLetter(j + 1) + currRow, cellData, styleIndex);
        break;
      default:
        row += addStringCell(self, getColumnLetter(j + 1) + currRow, cellData, styleIndex);
      }
    }

    row += '</x:row>';
    rows += row;
  }
  
  if (colsWidth !== '') {
    sheetFront += '<cols>' + colsWidth + '</cols>';
  }

  xlsx.file(config.fileName, sheetFront + '<x:sheetData>' + rows + '</x:sheetData>' + sheetBack);
};

module.exports = Sheet;

