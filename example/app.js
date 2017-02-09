const express = require('express');
const nodeExcel = require('../index');
const app = express();

app.get('/', function (req, res) {
  const conf = {
    stylesXmlFile: __dirname + '/styles.xml',
    name: 'report',
    cols: [
      {
        caption: '',
        type: 'string',
        width: 15
      },
      {
        caption: '',
        type: 'string',
        width: 15
      },
      {
        caption: '',
        type: 'string',
        width: 30
      }
    ],
    rows: [
      [{ value: '', style: 1 }, { value: '', style: 1 }, { value: '', style: 1 }],
      [{ value: '', style: 1 }, { value: 'THIS IS THE REPORT NAME', style: 1 }, { value: '', style: 1 }],
      [{ value: '', style: 1 }, { value: '', style: 1 }, { value: '', style: 1 }],
      ['Report Generated: 2017-02-08', '', ''],
      ['Event: Dat event do', '', ''],
      ['By: Francois Laubscher', '', ''],
      ['', '', ''],
      [{ value: 'First Name', style: 1 }, { value: 'Last Name', style: 1 }, { value: 'Email', style: 1 }],
      ['Bruce', 'Wayne', 'info@bat.man'],
      ['Clark', 'Kent', 'info@super.man'],
      ['Peter', 'Parker', 'info@spider.man']
    ]
  };

  const result = nodeExcel.execute(conf);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats');
  res.setHeader('Content-Disposition', 'attachment; filename=Report.xlsx');
  return res.send(new Buffer(result, 'binary'));
});

app.listen(3001, function () {
  console.log('listening on port 3001');
});