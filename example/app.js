var express = require('express');
var nodeExcel = require('../index');
var app = express();

app.get('/', function (req, res) {
  var conf = {
    stylesXmlFile: __dirname + '/styles.xml',
    name: 'report',
    cols: [
      {
        caption: 'First Name',
        type: 'string',
        width: 15
      },
      {
        caption: 'Last Name',
        type: 'string',
        width: 15
      },
      {
        caption: 'Email',
        type: 'string',
        width: 30
      }
    ],
    rows: [
      ['Bruce', 'Wayne', 'info@bat.man'],
      ['Clark', 'Kent', 'info@super.man'],
      ['Peter', 'Parker', 'info@spider.man']
    ]
  };

  var result = nodeExcel.execute(conf);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats');
  res.setHeader('Content-Disposition', 'attachment; filename=Report.xlsx');
  res.send(new Buffer(result, 'binary'));
});

app.listen(3001, function () {
  console.log('listening on port 3001');
});