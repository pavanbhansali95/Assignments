const express = require('express');
const bodyParser = require('body-parser');
const tempFile = require('tempfile');
const cors = require('cors');
const fs = require('fs');
const app = express();
const xl = require('excel4node');
const uuid = require('uuid');
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended: false}));
app.use(cors());
var json = [{student :'pavan',marks :'20'}];
function getExcel() {
    return new Promise((resolve, reject) => {
        const excel = require('exceljs');
        const workbook = new excel.Workbook();
    var worksheet = workbook.addWorksheet('mysheet');
    var keys = Object.keys(data[0]);
    const columns = [];
    workbook.creator = 'Me';
    workbook.lastModifiedBy = 'Her';
    workbook.created = new Date(1985, 8, 30);
    workbook.modified = new Date();
    workbook.lastPrinted = new Date(2016, 9, 27);
    workbook.properties.date1904 = true;

    workbook.views = [
        {
            x: 0, y: 0, width: 10000, height: 20000,
            firstSheet: 0, activeTab: 1, visibility: 'visible'
        }
    ];
    worksheet.columns = [
        { header: 'student', key: 'student', width: 10 },
        { header: 'marks', key: 'marks', width: 32 },
    ];
    worksheet.addRow({ student : 'rahul', marks: 20  });
    worksheet.addRow({ student : 'rah', marks: 23  });


    // keys.map((key) => {
    //     const headerObj = {
    //         header: key, key: key,width: 10
    //     };
    //     columns.push(headerObj);
    // });
    // worksheet.columns = columns;
    // data.map((tableData)=> {
    //     worksheet.addRow(tableData);
    // });
    const guidForClient = uuid.v1();
let pathNameWithGuid = `${guidForClient}_result.xlsx`;
workbook.xlsx.writeFile(pathNameWithGuid).then(() => {
        let stream = fs.createReadStream(pathNameWithGuid);
        stream.on("close", () => {
            fs.unlink(pathNameWithGuid, (error) => {
                if (error) {
                    throw error;
                }
            });
        });
        resolve(stream);
    },
    (err) => {
        throw err;
    }
    );
});
}
const createSheet = () => {

    return new Promise(resolve => {
  
  // setup workbook and sheet
  var wb = new xl.Workbook();
  
  var ws = wb.addWorksheet('Sheet');
  
  // Add a title row
  
  ws.cell(1, 1)
    .string('student')
  
  ws.cell(1, 2)
    .string('marks')
    for (let i = 0; i < json.length; i++) {

        let row = i + 2
      
        ws.cell(row, 1)
          .string('ppp')
      
        ws.cell(row, 2)
          .string('20')
      
      
      }
      
      resolve( wb )
      
        })
      }

app.get('/convertGridToExcel', function(req,res) {
//     const excel = require('exceljs');
//     const workbook = new excel.Workbook();
// var worksheet = workbook.addWorksheet('mysheet');
// var keys = Object.keys(data[0]);
// const columns = [];
// workbook.creator = 'Me';
// workbook.lastModifiedBy = 'Her';
// workbook.created = new Date(1985, 8, 30);
// workbook.modified = new Date();
// workbook.lastPrinted = new Date(2016, 9, 27);
// workbook.properties.date1904 = true;

// workbook.views = [
//     {
//         x: 0, y: 0, width: 10000, height: 20000,
//         firstSheet: 0, activeTab: 1, visibility: 'visible'
//     }
// ];
// worksheet.columns = [
//     { header: 'student', key: 'student', width: 10 },
//     { header: 'marks', key: 'marks', width: 32 },
// ];
// worksheet.addRow({ student : 'rahul', marks: 20  });
// worksheet.addRow({ student : 'rah', marks: 23  });

//     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//     res.setHeader("Content-Disposition", "attachment; filename=" + "Report.xlsx");
//     workbook.xlsx.write(res)
//         .then(function () {
//             res.end();
//             console.log('File write done........');
//         });
createSheet().then( file => {
    file.write('ExcelFile.xlsx', JSON.stringify(res));
  });

});
app.listen(8000, function () {
    console.log('Excel app listening on port 8000');
  });