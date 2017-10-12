// node_xj = require("xls-to-json");
//   node_xj({
//     input: "./Data/test.xlsx",  // input xls 
//     output: "output.json" // output json 
//   }, function(err, result) {
//     if(err) {
//       console.error(err);
//     } else {
//       //console.log(result.splice(0,1));
//       for(i=7;i<=result.length-1;i++){
//         console.log(result[i]);
//       }
//     }
//   });

// var xlsx = require('xlsx');
// var data = xlsx.readFile("./Data/test.xlsx");
// console.log(data);


// console.log(Object.keys(data).length)
// for(i=0;i<=Object.keys(data).length;i++){
//   console.log(data[i]);
// }

// var worksheet = data.Sheets;

// for (var row in worksheet) {
//   for (var col in row) {
//     console.log(worksheet[row][col]);
//   }
// }




// var XLSX = require('xlsx');
// var workbook = XLSX.readFile('./Data/test.xlsx');
// var sheet_name_list = workbook.SheetNames;
// sheet_name_list.forEach(function(y) {
//     var worksheet = workbook.Sheets[y];
//     var headers = {};
//     var data = [];
//     for(z in worksheet) {
//         if(z[0] === '!') continue;
//         //parse out the column, row, and value
//         var tt = 0;
//         for (var i = 0; i < z.length; i++) {
//             if (!isNaN(z[i])) {
//                 tt = i;
//                 break;
//             }
//         };
//         var col = z.substring(0,tt);
//         var row = parseInt(z.substring(tt));
//         var value = worksheet[z].v;

//         //store header names
//         if(row == 1 && value) {
//             headers[col] = value;
//             continue;
//         }

//         if(!data[row]) data[row]={};
//         data[row][headers[col]] = value;
//     }
//     //drop those first two rows which are empty
//     data.shift();
//     data.shift();
//     console.log(data);
// });




// var XLSX = require('xlsx');
// var workbook = XLSX.readFile('./Data/test.xlsx');
// var sheet_name_list = workbook.SheetNames;
// sheet_name_list.forEach(function(y) {
//     var worksheet = workbook.Sheets[y];
//     var headers = {};
//     var data = [];
//     for(z in worksheet) {
//         if(z[0] === '!') continue;
//         //parse out the column, row, and value
//         var col = z.substring(0,1);
//         var row = parseInt(z.substring(1));
//         var value = worksheet[z].v;

//         //store header names
//         if(row == 1) {
//             headers[col] = value;
//             continue;
//         }

//         if(!data[row]) data[row]={};
//         data[row][headers[col]] = value;
//     }
//     //drop those first two rows which are empty
//     // data.shift();
    
//     // data.shift();
//     // data.shift();
//     // data.shift();
//     console.log(data);
// });


// XLSX.writeFile(data, 'out.xlsx');

// if(typeof require !== 'undefined') {
//     console.log('hey');
const XLSX = require('xlsx');
// }
var workbook = XLSX.readFile('./Data/test.xlsx');
var sheet_name_list = workbook.SheetNames;
sheet_name_list.forEach(function(y) { /* iterate through sheets */
  var worksheet = workbook.Sheets[y];
  for (z in worksheet) {

    //console.log(z);
    //console.log(JSON.stringify(worksheet[z]));

    console.log(JSON.stringify(worksheet[z]));

     for(i=0;i<=6 ;i++){
       
     }
    
    if(z[0] === '!') continue;
      
     //console.log(y + "!" + z + "=" + JSON.stringify(worksheet[z].v));
  }
  //console.log(worksheet);

});


XLSX.writeFile(workbook, 'out.xlsx');