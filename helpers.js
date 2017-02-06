exports.startTag = function (obj, tagName, closed) {
  var result = '<' + tagName;
  for (var p in obj) {
    result += ' ' + p + '=' + obj[p];
  }

  if (!closed) {
    result += '>';
  } else {
    result += '/>';
  }

  return result;
};

exports.endTag = function (tagName) {
  return `</${tagName}>`;
};

exports.addNumberCell = function (cellRef, value, styleIndex) {
  styleIndex = styleIndex || 0;
  if (value === null) {    
    return '';
  } else {
    return `<x:c r="${cellRef}" s="${styleIndex}" t="n"><x:v>${value}</x:v></x:c>`;
  }
};

exports.addDateCell = function (cellRef, value, styleIndex) {
  styleIndex = styleIndex || 1;
  if (value === null) {
    return '';
  } else {
    return `<x:c r="${cellRef}" s="${styleIndex}" t="n"><x:v>${value}</x:v></x:c>`;
  }
};

exports.addBoolCell = function (cellRef, value, styleIndex) {
  styleIndex = styleIndex || 0;
  if (value === null) {
    return '';
  }

  if (value) {
    value = 1;
  } else {
    value = 0;
  }

  return `<x:c r="${cellRef}" s="${styleIndex}" t="b"><x:v>${value}</x:v></x:c>`;
};


exports.addStringCell = function (sheet, cellRef, value, styleIndex) {
  styleIndex = styleIndex || 0;
  
  if (value === null) {
    return '';
  }

  if (typeof value === 'string') {
    value = value.replace(/&/g, '&amp;').replace(/'/g, '&apos;').replace(/>/g, '&gt;').replace(/</g, '&lt;');
  }

  return `<x:c r="${cellRef}" s="${styleIndex}" t="inlineStr"><x:is><x:t>${value}</x:t></x:is></x:c>`;
};

exports.getColumnLetter = function (col) {
  if (col <= 0) {
    throw 'col must be more than 0';
  }

  var array = [];
  while (col > 0) {
    var remainder = col % 26;
    col /= 26;
    col = Math.floor(col);
    if (remainder === 0) {
      remainder = 26;
      col--;
    }
    array.push(64 + remainder);
  }

  return String.fromCharCode.apply(null, array.reverse());
};
