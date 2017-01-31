# excel-export

A simple node.js module for exporting data set to Excel xlsx file.

## Using excel-export

A configs object is required to be passed in the `execute` method.
If generating multiple sheets, the configs object should be an array of multiple config objects.
If generating a single sheet, the config object should be a single object.

Example of a single worksheet
```javascript
{
  name: 'name.xlsx',
  cols: [
    {
      caption: 'column display name',
      type: 'string',
      width: 10,
      beforeCellWrite: function(row, cellData, option) {
        // do something with column data
      }
    }
  ],
  rows: [].
  stylesXmlFile: 'styles.xml'
}
```

- name: Specify worksheet name
- cols: `Array` of column definitions
  - caption: Column caption
  - type: string / date / bool / number
  - width: (optional) total characters in cell (blame excel)
  - beforeCellWrite: (optional) This function is invoked with row, cell data and option object. The return value is written to the cell.
- rows: `Array` of data to be exported. Data needs to be a 2D-`Array`. Each row should be the same length as cols.
- stylesXmlFile: Absolute path to Excel `styles.xml` file. An easy way to get a `styles.xml` file is to unzip an existing xlsx file which has the desired styles and copy the `styles.xml` file.

## Example usage with Express

```javascript
const express = require('express');
const nodeExcel = require('excel-export');
const app = express();

app.get('/', function (req, res) {
  const conf = {
    stylesXmlFile: __dirname + '/styles.xml',
    name: 'report',
    cols: [
      {
        caption: 'First Name',
        type: 'string',
        width: 250
      },
      {
        caption: 'Last Name',
        type: 'string',
        width: 250
      },
      {
        caption: 'Email',
        type: 'string',
        width: 250
      }
    ],
    rows: [
      ['Name 1', 'Name 2', 'Name 3'],
      ['Surname 1', 'Surname 2', 'Surname 3'],
      ['example@email.com', 'anotherexample@email.com', 'yetanotherexample@email.com']
    ]
  };

  const result = nodeExcel.execute(conf);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats');
  res.setHeader('Content-Disposition', 'attachment; filename=Report.xlsx');
  res.send(new Buffer(result, 'binary'));
});

app.listen(3001);
```
