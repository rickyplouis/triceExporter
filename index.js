const converter = require('node-xlsx');
const fs = require('fs');
const obj = converter.parse('./concrete.xlsx');

const formatHeader = (header) => {
  let proposedWidthCounter = 0;
  let proposedLengthCounter = 0;
  let proposedDepthCounter = 0;
  let actualWidthCounter = 0;
  let actualLengthCounter = 0;
  let actualDepthCounter = 0;
  return header.map((item) => {
    item = item.toLowerCase();
    if (item.indexOf('date completed') >= 0) {
      item = 'completionDate'
    }
    if (item.indexOf('order') >= 0) {
      item = 'orderNumber'
    }
    if (item.indexOf('release') >= 0) {
      item = 'releaseDate'
    }
    if (item.indexOf('c&g') >= 0) {
      item = 'proposedCG'
    }
    if (item.indexOf('proposed width') >= 0) {
      item = `proposedWidth_${proposedWidthCounter}`;
      proposedWidthCounter += 1;
    }
    if (item.indexOf('proposed length') >= 0) {
      item = `proposedLength_${proposedLengthCounter}`;
      proposedLengthCounter += 1;
    }
    if (item.indexOf('proposed depth') >= 0) {
      item = `proposedLength_${proposedDepthCounter}`;
      proposedDepthCounter += 1;
    }
    if (item.indexOf('actual width') >= 0) {
      item = `actualWidth_${actualWidthCounter}`;
      actualWidthCounter += 1;
    }
    if (item.indexOf('actual length') >= 0) {
      item = `actualLength_${actualLengthCounter}`;
      actualLengthCounter += 1;
    }
    if (item.indexOf('actual depth') >= 0) {
      item = `actualLength_${actualDepthCounter}`;
      actualDepthCounter += 1;
    }
    if (item.indexOf('dig') >= 0) {
      item = 'digNumber';
    }
    return item.replace(' ', '');
  });
}

const formatName = (name) => name.toLowerCase().replace(' ', '');

const createCSV = (sheets) => {
  //looping through all sheets
  let csvStringArray = [];
  console.log('sheets.length', sheets.length);
  for(let x = 0; x < sheets.length; x++) {
      let rows = [];
      let writeStr = "";

      let sheet = sheets[x];

      //loop through all rows in the sheet
      for(var y = 0; y < sheet['data'].length; y++) {
          //add the row to the rows array
          if (y === 0) {
            rows.push(formatHeader(sheet['data'][y]))
          } else {

            rows.push((sheet['data'][y]))
          }
      }
      //creates the csv string to write it to a file
      for(var z = 0; z < rows.length; z++) {
        writeStr += rows[z].join(",") + "\n";
      }
      //writes to a file, but you will presumably send the csv as a
      //response instead
      csvStringArray.push({name: formatName(sheet.name), data: writeStr});
  }
  csvStringArray.map((str) => {
    const dir = './outputCSV';
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir);
    }
    fs.writeFile(`./outputCSV/${str.name}.csv`, str.data, function(err) {
      console.log('writing', str.name);
      if(err) {
        return console.log(err);
      }
      console.log('saved!');
    });
  })
}

createCSV(obj);
