
// var xlsx = require('xlsx');

// var workbook = xlsx.readFile('data.xlsm');
// var workbook2 = xlsx.readFile('test1.xlsx');



// var worksheet = workbook.Sheets[workbook.SheetNames[0]];

// // console.log(worksheet);

// var worksheet2 = workbook2.Sheets[workbook2.SheetNames[0]];

// console.log(worksheet2);

// console.log(worksheet['A1']);
// console.log(worksheet2['A2']);


// var _MS_PER_DAY = 1000 * 60 * 60 * 24;

// a and b are javascript Date objects
// function dateDiffInDays(b, a) {
//   // Discard the time and time-zone information.
//   var utc1 = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
//   var utc2 = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());

//   return Math.floor((utc2 - utc1) / _MS_PER_DAY);
// }

// // test it
// var a = new Date("2017-01-01"),
//     b = new Date("2017-01-25"),
//     difference = dateDiffInDays(b, a);
// console.log(difference);




// var alpha = ['A']//,'B','C','D','E','F','G','H','I','J','K','L','M','N','O'];

// var num = ['1','2','3','4','5','6','7','8','9','10'];

// console.log(xlsx.utils.encode_cell({c:0,r:0}))

// l=0;

// for( y=8;;y=y+2){
//     if(worksheet['A'+y] == undefined || worksheet['A'+y].v==''){
//         break;
//     }
//     l++;
// }

// console.log(worksheet["B8"].w);
// console.log(worksheet["D8"].w);




// console.log(worksheet['A1']);

// num.forEach(n1 => {
//     alpha.forEach(a1=>{
//         console.log(worksheet[a1+n1]);
//     })
//  });

