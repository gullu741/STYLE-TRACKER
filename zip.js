// const jszip = require("jszip");

// var zip = new jszip();

// zip.file("./hello.txt",'hello WQorld\n');

var unzip = require("unzip")
var fs = require("fs");
var JSZip = require("jszip");
var parseString = require('xml2js').parseString;

var zip = new JSZip();
// zip.file("file", content);
// ... and other manipulations
// zip
// .generateNodeStream({type:'nodebuffer',streamFiles:true})
// .pipe(fs.createWriteStream('out.zip'))
// .on('finish', function () {
//     // JSZip generates a readable stream with a "end" event,
//     // but is piped here in a writable stream which emits a "finish" event.
//     console.log("out.zip written.");
// });

function formattxt(name) {
    n = name.indexOf(':');
    if (n != -1) {
        name = name.slice(n + 1, name.length);
    }
    return name;
}


// fs.readFile("D:/node/APP/test1/xl/drawings/drawing1.xml", (err, data) => {
//     if (err) {
//         console.log(err);
//     }

//     parseString(data, {
//         tagNameProcessors: [formattxt],
//         attrNameProcessors:[formattxt]
//     }
//         , function (err, data) {
//             if (err) console.log(err);
//             data = data.wsDr.twoCellAnchor
//             console.log(data);

//         })
// })
// console.log("Reading file");
// fs.readFile("./test1.xlsx", (err, fdata) => {
//     if (err) {
//         console.log(err);
//     }
//     console.log(fdata);
//     console.log("File Read Sucessfully");

//     JSZip.loadAsync(fdata).then(function (zdata) {
//         //  zdata = zdata.files
//         //  zdata = zdata["xl/drawings/drawing1.xml"]
//         // zdata = zdata.files["xl/drawings/drawing1.xml"]["_data"]["compressedContent"];
//         zdata = zdata.file("xl/drawings/drawing1.xml");
//         console.log(zdata);
//             parseString(zdata, function (err3, data) {
//                 if (err3) {
//                     console.log(err3);
//                     return;
//                 }
//                 console.log("done converting data");
//                 data = data.wsDr.twoCellAnchor
//                 console.log(data);
//                 console.log("DONE");

//             })
//     })

// })

// fs.createReadStream("./test1.xlsx").pipe(unzip.Extract({path:"./"}));
// var readStream = fs.createReadStream('./test1.xlsx');
// var writeStream = fs.Writer('./test1');

// readStream
//   .pipe(unzip.Parse())
//   .pipe(writeStream)


const admZip = require("adm-zip");
zip = new admZip('D:/node/APP/data.xlsm');
zip.extractAllTo(`D:/node/APP/xldata`, true)
// console.log(zip)
var images = [];
fs.readFile("D:/node/APP/xldata/xl/drawings/drawing1.xml", (err, data) => {
    if (err) {
        console.log(err);
    }

    parseString(data, {
        tagNameProcessors: [formattxt],
        attrNameProcessors: [formattxt]
    }
        , function (err, data) {
            if (err) console.log(err);
            data = data.wsDr.twoCellAnchor
            
            data.forEach(function(datas){
                console.log(datas['pic'][0].blipFill[0].blip[0]['$'].embed.split('d')[1])
            })
            
            
            // data.forEach(function (idata) {
            //     if(idata.from[0].col[0]==9){
            //         images.push({
            //             row:idata.from[0].row[0],
            //             iname:idata.pic[0].nvPicPr[0].cNvPr[0]['$'].name
            //         })
            //         // console.log(idata.from[0].row[0])
            //     }
            //     // console.log(idata.from[0].row[0])
            //     // console.log(idata.pic[0].nvPicPr[0].cNvPr[0]['$'].name);
            // })
            // console.log(data[0].pic[0].nvPicPr[0].cNvPr[0]['$'].name);
            // console.log(images)
        })
})

