

// import { isUndefined } from 'util';

//  $= jQuery = require('jquery');
// require("jquery-ui");
const xlsx = require('xlsx');
const fs = require('fs')
const { ipcRenderer } = require('electron');
const { execSync } = require('child_process');
const admZip = require('adm-zip');

var parseString = require('xml2js').parseString;

function validData(data) {
    try {
        return data.w;
    }
    catch (err) {
        return " ";
    }
}

function formattxt(name) {
    n = name.indexOf(':');
    if (n != -1) {
        name = name.slice(n + 1, name.length);
    }
    return name;
}
var _MS_PER_DAY = 1000 * 60 * 60 * 24;

// a and b are javascript Date objects
function dateDiffInDays(b, a) {
    // Discard the time and time-zone information.
    var utc1 = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
    var utc2 = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());

    return Math.floor((utc2 - utc1) / _MS_PER_DAY);
}




function openfile() {
    console.log("sending event");
    ipcRenderer.send('openFile', () => {
        console.log("Event sent to main process");
    })
}

ipcRenderer.on('filename', (event, data) => {
    createNewTable(data);
})
var images = [];

function readdata(filename) {
    wb = xlsx.readFile(String(filename));
    if (wb == undefined) {
        console.log("File read error");
    }
    else {
        images = [];
        console.log('workbook loaded suscessfully')
        console.log(String(filename))
        ws = wb.Sheets[wb.SheetNames[0]];

        console.log('sheet loaded suscess fully');
        if (ws['A1'].w.toUpperCase() != 'STYLE TRACKER') {
            alert("Please select a valid file with required datastructure");
            return;
        }

        zip = new admZip(String(filename));
        // execSync(`mkdir ${__dirname}\\xldata`);
        zip.extractAllTo(`./xldata`, true);

        data = fs.readFileSync(`./xldata/xl/drawings/drawing1.xml`)

        parseString(data, { tagNameProcessors: [formattxt], attrNameProcessors: [formattxt] }, function (err, data) {
            if (err) console.log(err)
            data = data.wsDr.twoCellAnchor;
            fs.writeFile("./data.json",JSON.stringify(data),function(err){
                if(err)console.log(err)
                console.log('file written');
            })
            data.forEach(function (idata) {
                if (idata.from[0].col[0] == 9) {
                    images[idata.from[0].row[0]] = idata['pic'][0].blipFill[0].blip[0]['$'].embed.split('d')[1];
                    // images.push({
                    //     row: idata.from[0].row[0] + 1,
                    //     iname: idata.pic[0].nvPicPr[0].cNvPr[0]['$'].name
                    // })
                }

            })
            console.log(images);
        })






        $('.datatable').empty();
        $("#passed").empty();
        $("#approaching").empty();
        $('#ndate').empty();
        $('#vname').empty();
        console.log(__dirname);
        console.log('./')

        console.log('emptying div');
        //First table based on m 
        content = '<table class = "table ">';

        content += '<thead><tr>' +
            `<th>${validData(ws['A7'])}</th>` +// a7 
            `<th>${validData(ws['B7'])}</th>` +// b7 
            `<th>${validData(ws['E7'])}</th>` +// e7 
            `<th>${validData(ws['I7'])}</th>` +// i7 
            // `<th>${validData(ws['F7'])}</th>` +// f7 
            `<th>${validData(ws['J7'])}</th>` +// j7         
            `<th>${validData(ws['K7'])}</th>` +// k7 
            `<th>${validData(ws['M7'])}</th>` +// m7 
            // `<th>${validData(ws['O7'])}</th>` +// o7 
            '</tr></thead>'
        content += '<tbody>';
        danger = 0;
        warning = 0;
        for (i = 0; ; i++) {
            // console.log(ws['A' + (8 + i * 2)]);
            if (ws['A' + (8 + i * 2)] == undefined) {
                break;
            }
            else {
                if (ws['M' + (9 + i * 2)] == undefined) {
                    console.log(ws['M' + 9 + i * 2])
                    diff = dateDiffInDays(new Date(), new Date(validData(ws['M' + (8 + i * 2)])));
                    console.log(diff);
                    if (diff >= 0) {
                        content += '<tr bgcolor="#FF0000">';
                        danger++;
                    } else if (diff < 0 && diff >= -3) {
                        content += '<tr bgcolor = "#06B9CF">';
                        warning++;
                    } else if (diff < -3 && diff >= -7) {
                        content += '<tr bgcolor = "#FDFD96">';
                        warning++;
                    }
                    else {
                        continue;
                    }



                    content += `<td>${validData(ws['A' + (8 + i * 2)])}</td>` +// a7 
                        `<td>${validData(ws['B' + (8 + i * 2)])}</td>` +// b7 
                        `<td>${validData(ws['E' + (8 + i * 2)])}</td>` +// e7
                        `<td>${validData(ws['I' + (8 + i * 2)])}</td>` +// i7 
                        // `<td>${
                        // (new Date(validData(ws['D' + (8 + i * 2)])) == 'Invalid Date') ? " " : new Date(validData(ws['D' + (8 + i * 2)])).toLocaleDateString()
                        // }</td>` +// d7 
                        // `<td>${
                        // (new Date(validData(ws['F' + (8 + i * 2)])) == 'Invalid Date') ? " " : new Date(validData(ws['F' + (8 + i * 2)])).toLocaleDateString()
                        // }</td>` +// f7 
                        `<td>${validData(ws['J' + (8 + i * 2)])}<img src= "../../xldata/xl/media/image${images[(8 + i * 2)-1]}.png" style="width:100px;height:100px;"></td>` +// j7         
                        `<td>${validData(ws['K' + (8 + i * 2)])}</td>` +// k7 
                        `<td>${
                        (new Date(validData(ws['M' + (8 + i * 2)])) == 'Invalid Date') ? " " : new Date(validData(ws['M' + (8 + i * 2)])).toLocaleDateString()
                        }</td>` +// m7 
                        // `<td>${
                        // (new Date(validData(ws['O' + (8 + i * 2)])) == 'Invalid Date') ? " " : new Date(validData(ws['O' + (8 + i * 2)])).toLocaleDateString()
                        // }</td>` +// o7 
                        '</tr>'
                
                    console.log("image addrress = " + images[(8 + i * 2)-1]);
                }
            }
        }



        content += '</tbody></table>';
        // console.log(content)
        // Second table based on O
        content += '<table class = "table ">';

        content += '<thead><tr>' +
            `<th>${validData(ws['A7'])}</th>` +// a7 
            `<th>${validData(ws['B7'])}</th>` +// b7 
            `<th>${validData(ws['E7'])}</th>` +// c7 
            `<th>${validData(ws['I7'])}</th>` +// d7 
            // `<th>${validData(ws['F7'])}</th>` +// f7 
            `<th>${validData(ws['J7'])}</th>` +// j7         
            `<th>${validData(ws['K7'])}</th>` +// k7 
            // `<th>${validData(ws['M7'])}</th>` +// m7 
            `<th>${validData(ws['O7'])}</th>` +// o7 
            '</tr></thead>'
        content += '<tbody>';
        // danger = 0;
        // warning = 0;
        for (i = 0; ; i++) {
            // console.log(ws['A' + (8 + i * 2)]);
            if (ws['A' + (8 + i * 2)] == undefined) {
                break;
            }
            else {
                if (ws['O' + (9 + i * 2)] == undefined) {
                    diff = dateDiffInDays(new Date(), new Date(validData(ws['O' + (8 + i * 2)])));
                    console.log(diff);
                    if (diff >= 0) {
                        content += '<tr bgcolor="#FF0000">';
                        danger++;
                    } else if (diff < 0 && diff >= -3) {
                        content += '<tr bgcolor = "#06B9CF">';
                        warning++;
                    } else if (diff < -3 && diff >= -7) {
                        content += '<tr bgcolor = "#FDFD96">';
                        warning++;
                    } else {
                        continue;
                    }



                    content += `<td>${validData(ws['A' + (8 + i * 2)])}</td>` +// a7 
                        `<td>${validData(ws['B' + (8 + i * 2)])}</td>` +// b7 
                        `<td>${validData(ws['E' + (8 + i * 2)])}</td>` +// c7 
                        `<td>${validData(ws['I' + (8 + i * 2)])}</td>` +// c7 
                        `<td>${validData(ws['J' + (8 + i * 2)])}<img src= "../../xldata/xl/media/image${images[(8 + i * 2)-1]}.png" style="width:100px;height:100px;"></td>` +// j7         
                        `<td>${validData(ws['K' + (8 + i * 2)])}</td>` +// k7 
                        `<td>${
                        (new Date(validData(ws['O' + (8 + i * 2)])) == 'Invalid Date') ? " " : new Date(validData(ws['O' + (8 + i * 2)])).toLocaleDateString()
                        }</td>` +// o7 
                        '</tr>'
                }
            }
        }



        content += '</tbody></table>';

        // third table based on rm 
        content += '<table class = "table ">';

        content += '<thead><tr>' +
            `<th>${validData(ws['A7'])}</th>` +// a7 
            `<th>${validData(ws['B7'])}</th>` +// b7 
            `<th>${validData(ws['E7'])}</th>` +// e7 
            `<th>${validData(ws['R7'])}</th>` +// r7 
            // `<th>${validData(ws['F7'])}</th>` +// f7 
            `<th>${validData(ws['S7'])}</th>` +// s7         
            // `<th>${validData(ws['K7'])}</th>` +// k7 
            // `<th>${validData(ws['M7'])}</th>` +// m7 
            // `<th>${validData(ws['O7'])}</th>` +// o7 
            '</tr></thead>'
        content += '<tbody>';
        // danger = 0;
        // warning = 0;
        for (i = 0; ; i++) {
            // console.log(ws['A' + (8 + i * 2)]);
            if (ws['A' + (8 + i * 2)] == undefined) {
                break;
            }
            else {
                if (ws['S' + (9 + i * 2)] == undefined && ws['R' + (8 + i * 2)] != undefined ) {
                    diff = dateDiffInDays(new Date(), new Date(validData(ws['S' + (8 + i * 2)])));
                    console.log(diff);
                    if (diff >= 0) {
                        content += '<tr bgcolor="#FF0000">';
                        danger++;
                    } else if (diff < 0 && diff >= -3) {
                        content += '<tr bgcolor = "#06B9CF">';
                        warning++;
                    } else if (diff < -3 && diff >= -7) {
                        content += '<tr bgcolor = "#FDFD96">';
                        warning++;
                    } else {
                        continue;
                    }



                    content += `<td>${validData(ws['A' + (8 + i * 2)])}</td>` +// a7 
                        `<td>${validData(ws['B' + (8 + i * 2)])}</td>` +// b7 
                        `<td>${validData(ws['E' + (8 + i * 2)])}</td>` +// e7 
                        `<td>${validData(ws['R' + (8 + i * 2)])}</td>` +// r7 
                        // `<td>${validData(ws['J' + (8 + i * 2)])}</td>` +// j7         
                        // `<td>${validData(ws['K' + (8 + i * 2)])}</td>` +// k7 
                        `<td>${
                        (new Date(validData(ws['S' + (8 + i * 2)])) == 'Invalid Date') ? " " : new Date(validData(ws['S' + (8 + i * 2)])).toLocaleDateString()
                        }</td>` +// o7 
                        '</tr>'
                }
            }
        }



        content += '</tbody></table>';
        //   fourth table based on trims
        content += '<table class = "table ">';

        content += '<thead><tr>' +
            `<th>${validData(ws['A7'])}</th>` +// a7 
            `<th>${validData(ws['B7'])}</th>` +// b7 
            `<th>${validData(ws['U7'])}</th>` +// c7 
            `<th>${validData(ws['V7'])}</th>` +// d7 
            // `<th>${validData(ws['F7'])}</th>` +// f7 
            // `<th>${validData(ws['J7'])}</th>` +// j7
            // `<th>${validData(ws['K7'])}</th>` +// k7 
            // `<th>${validData(ws['M7'])}</th>` +// m7 
            // `<th>${validData(ws['O7'])}</th>` +// o7 
            '</tr></thead>'
        content += '<tbody>';
        // danger = 0;
        // warning = 0;
        for (i = 0; ; i++) {
            // console.log(ws['A' + (8 + i * 2)]);
            if (ws['A' + (8 + i * 2)] == undefined) {
                break;
            }
            else {
                if (ws['V' + (9 + i * 2)] == undefined && ws['U' + (8 + i * 2)] != undefined) {
                    diff = dateDiffInDays(new Date(), new Date(validData(ws['V' + (8 + i * 2)])));
                    console.log(diff);
                    if (diff >= 0) {
                        content += '<tr bgcolor="#FF0000">';
                        danger++;
                    } else if (diff < 0 && diff >= -3) {
                        content += '<tr bgcolor = "#06B9CF">';
                        warning++;
                    } else if (diff < -3 && diff >= -7) {
                        content += '<tr bgcolor = "#FDFD96">';
                        warning++;
                    } else {
                        continue;
                    }



                    content += `<td>${validData(ws['A' + (8 + i * 2)])}</td>` +// a7 
                        `<td>${validData(ws['B' + (8 + i * 2)])}</td>` +// b7 
                        `<td>${validData(ws['U' + (8 + i * 2)])}</td>` +// c7 
                        // `<td>${validData(ws['I' + (8 + i * 2)])}</td>` +// c7 
                        // `<td>${validData(ws['J' + (8 + i * 2)])}</td>` +// j7         
                        // `<td>${validData(ws['K' + (8 + i * 2)])}</td>` +// k7 
                        `<td>${
                        (new Date(validData(ws['V' + (8 + i * 2)])) == 'Invalid Date') ? " " : new Date(validData(ws['V' + (8 + i * 2)])).toLocaleDateString()
                        }</td>` +// o7 
                        '</tr>'
                }
            }
        }



        content += '</tbody></table>';

        $('.datatable').append(content);
        $('#passed').append("Passed: " + danger);
        $('#approaching').append("Approaching: " + warning);
        $('#ndate').append(new Date().toLocaleDateString());
        $('#vname').append(validData(ws['B2']));

    }
}


function createNewTable(filename) {
    if (filename == undefined) {
        return;
    }
    else {

        readdata(filename);
    }

}

$('document').ready(() => {
    $('#cdate').datepicker();
    $('#odate').on('click', function () {
        $(this).datepicker('option', 'showAnim', 'slideDown')
    })
});


// function write(){
//     document.write(`<table style="width:100%">
//       <tr>
//         <th>Firstname</th>  
//         <th>Lastname</th> 
//         <th>Age</th>
//     </tr>
//     <tr>
//         <td>Jill</td>
//         <td>Smith</td> 
//         <td>50</td>
//     </tr>
//     <tr>
//         <td>Eve</td>
//         <td>Jackson</td> 
//         <td>94</td>
//     </tr>
//     </table>`);
// }