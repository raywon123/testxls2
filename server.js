var express = require('express');
var excel = require('exceljs');
var app = express();

app.use(express.static('public'));

// route:
app.get('/download', (req, res) => res.download('./sample.pdf'));

//----------- generate xls
app.get('/download2', (req, res) => {

    
    var workbook = new excel.Workbook(); //creating workbook
    var sheet = workbook.addWorksheet('MySheet'); //creating worksheet

    var image = workbook.addImage({
        filename: './pngImage.png',
        extension: 'png'
    }); // adding an image in workbook first
    sheet.addImage(image, 'D10:G14'); // adding an image in the worksheet in particular place

    var objArray = [{
        "id": 0,
        "name": "xxxx",
        "is_active": "false"
    }, {
        "id": 1,
        "name": "yyyy",
        "is_active": "true"
    }]
    sheet.addRow().values = Object.keys(objArray[0]);

    objArray.forEach(function (item) {
        var valueArray = [];
        valueArray = Object.values(item); // forming an array of values of single json in an array
        sheet.addRow().values = valueArray; // add the array as a row in sheet
    })

    workbook.xlsx.writeFile('./sample.xlsx').then(function () {
        console.log("file is written");

        // res.download('./sample.xlsx');

        res.download('./sample.xlsx', 'sample_user_facing.xlsx', (err) => {
            if (err) {
              //handle error
              console.log("send fail")
              return
            } else {
              //do something
              console.log("send success")
            }
          })
    })

    
});
//-----------

console.log("server is up at http://localhost:3000");
app.listen(3000);