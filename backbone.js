var app = require('electron').remote;
var dialog = app.dialog;
var fs = require('fs');
var Excel = require('exceljs');
var parse = require('csv-parse/lib/sync');
var tTData = {};
var tCData = {};
var tC = [];
var coolingData = [];
var heatingData = [];

document.getElementById('T-select-file').addEventListener('click', function() {
    dialog.showOpenDialog(function(fileNames) {
        if (fileNames === undefined) {
            console.log("No file selected");
        } else {
            document.getElementById("T-actual-file").value = fileNames[0];
            var contentEditor = document.getElementById("T-content-editor");
            readFile(fileNames[0], contentEditor, tTData);
        }
    });
}, false);

document.getElementById('C-select-file').addEventListener('click', function() {
    dialog.showOpenDialog(function(fileNames) {
        if (fileNames === undefined) {
            console.log("No file selected");
        } else {
            document.getElementById("C-actual-file").value = fileNames[0];
            var contentEditor = document.getElementById("C-content-editor");
            readFile(fileNames[0], contentEditor, tCData);
        }
    });
}, false);

document.getElementById('convertButton').addEventListener('click', function() {
    if (tTData.data && tCData.data) {
        convert();
        dialog.showSaveDialog(function(actualFilePath) {
            if(!(/\.xlsx$/.test(actualFilePath))) {
                actualFilePath = actualFilePath + '.xlsx';
            }
            saveChanges(actualFilePath);
        });
    } else {
        alert('Please import the experiment data files first.');
    }
}, false);

function readFile(filepath, contentEditor, result) {
    fs.readFile(filepath, 'utf-8', function(err, data) {
        if (err) {
            alert("An error ocurred reading the file :" + err.message);
            return;
        }
        contentEditor.value = data;
        result.data = data;
    });
}

function saveChanges(filepath) {
    var workbook = new Excel.Workbook();
    workbook.created = new Date();
    var coolingDataSheet = workbook.addWorksheet('Cooling Curve');
    coolingDataSheet.columns = [
        { header: 'T', key: 'T', width: 10 },
        { header: 'C*', key: 'C', width: 10 },
        { header: 'L*', key: 'L', width: 10 },
        { header: 'a*', key: 'a', width: 10 },
        { header: 'b*', key: 'b', width: 10 },
        { header: 'h°', key: 'h', width: 10 }
    ];
    var length = coolingData.length;
    for(var i=0; i<length; ++i) {
        coolingDataSheet.addRow({
            T: roundToTwoDigit(coolingData[i].T),
            C: coolingData[i].C,
            L: coolingData[i].L,
            a: coolingData[i].a,
            b: coolingData[i].b,
            h: coolingData[i].h
        });
    }
    var heatingDataSheet = workbook.addWorksheet('Heating Curve');
    heatingDataSheet.columns = coolingDataSheet.columns;
    length = heatingData.length;
    for(var i=0; i<length; ++i) {
        heatingDataSheet.addRow({
            T: roundToTwoDigit(heatingData[i].T),
            C: heatingData[i].C,
            L: heatingData[i].L,
            a: heatingData[i].a,
            b: heatingData[i].b,
            h: heatingData[i].h
        });
    }
    workbook.xlsx.writeFile(filepath);
}

function roundToTwoDigit(a) {
    return Math.round(a*100)/100;
}

// Main data processing logic
function convert() {
    var tT = [];
    var tTOutput = parse(tTData.data, {
        columns: true,
        ltrim: true,
        auto_parse: true
    });
    var arrayLength = tTOutput.length;
    for (let i = 0; i < arrayLength; i++) {
        tT.push({
            timestampMs: Date.parse(tTOutput[i].Date + ' ' + tTOutput[i].Time) / 1000,
            P: tTOutput[i].P,
            SV: tTOutput[i].SV,
        });
    }

    var tCOutput = parse(tCData.data, {
        columns: true,
        ltrim: true,
        auto_parse: true
    });
    arrayLength = tCOutput.length;
    var timeName= arrayLength > 0 && 'Trial Name' in tCOutput[0] ? 'Trial Name' : 'Name';
    for (let i = 0; i < arrayLength; i++) {
        tC.push({
            timestampMs: Date.parse(tCOutput[i][timeName].split('@')[1]) / 1000,
            C: tCOutput[i]['C* '],
            L: tCOutput[i]['L* '],
            a: tCOutput[i]['a* '],
            b: tCOutput[i]['b* '],
            h: tCOutput[i]['h° '],
        });
    }

    var merged = tT.concat(tC);
    merged.sort(function(a, b) {
        if (a.timestampMs == b.timestampMs) return 0;
        return a.timestampMs > b.timestampMs ? 1 : -1;
    });

    arrayLength = merged.length;
    let prev;
    let next;
    next = 0;
    while (next < arrayLength) {
        if (merged[next].P != null) {
            if (prev == null) {
                for (let i = 0; i < next; i++) {
                    merged[i].T = merged[next].P;
                }
            } else {
                for (let i = prev + 1; i < next; i++) {
                    it = merged[i].timestampMs;
                    prevt = merged[prev].timestampMs;
                    nextt = merged[next].timestampMs;
                    if(prevt == nextt) {
                        merged[i].T = (merged[prev].P + merged[next].P) / 2;
                    } else {
                        merged[i].T = (merged[next].P * (it - prevt) + merged[prev].P * (nextt - it)) / (nextt - prevt);
                    }
                }
            }
            prev = next;
        }
        next++;
    }
    for (let i = prev + 1; i < arrayLength; i++) {
        merged[i].T = merged[prev].P;
    }
    var minIndex = 0;
    var tcLength = tC.length;
    for(let i=0; i<tcLength; ++i) {
        if(tC[minIndex].T>tC[i].T) {
            minIndex = i;
        }
    }
    coolingData = tC.slice(0, minIndex+1);
    heatingData = tC.slice(minIndex);
}
