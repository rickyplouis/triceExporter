const converter = require('node-xlsx');
const fs = require('fs');

const csv = require('csvtojson');
const moment = require('moment');
const readline = require('readline');
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

let userInput = '';


const ExcelDateToJSDate = (serial) => {
   var utc_days  = Math.floor(serial - 25569);
   var utc_value = utc_days * 86400;
   var date_info = new Date(utc_value * 1000);
   var fractional_day = serial - Math.floor(serial) + 0.0000001;
   var total_seconds = Math.floor(86400 * fractional_day);
   var seconds = total_seconds % 60;
   total_seconds -= seconds;
   var hours = Math.floor(total_seconds / (60 * 60));
   var minutes = Math.floor(total_seconds / 60) % 60;
   return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
}

const formatHeader = (header) => {
  let proposedWidthCounter = 0;
  let proposedLengthCounter = 0;
  let proposedDepthCounter = 0;
  let actualWidthCounter = 0;
  let actualLengthCounter = 0;
  let actualDepthCounter = 0;
  return header.map((item) => {
    if (item.indexOf('Date Completed') >= 0) {
      item = 'completionDate'
    }
    if (item.indexOf('Order') >= 0) {
      item = 'orderNumber'
    }
    if (item.indexOf('Release') >= 0) {
      item = 'releaseDate'
    }
    if (item.indexOf('Address') >= 0) {
      item = 'address'
    }
    if (item.indexOf('Town') >= 0) {
      item = 'town'
    }
    if (item.indexOf('Area') >= 0) {
      item = 'area'
    }
    if (item.indexOf('C&G') >= 0) {
      item = 'proposedCG'
    }
    if (item.indexOf('Proposed Width') >= 0) {
      item = `proposedWidth_${proposedWidthCounter}`;
      proposedWidthCounter += 1;
    }
    if (item.indexOf('Proposed Length') >= 0) {
      item = `proposedLength_${proposedLengthCounter}`;
      proposedLengthCounter += 1;
    }
    if (item.indexOf('Proposed Depth') >= 0) {
      item = `proposedDepth_${proposedDepthCounter}`;
      proposedDepthCounter += 1;
    }
    if (item.indexOf('Actual Width') >= 0) {
      item = `actualWidth_${actualWidthCounter}`;
      actualWidthCounter += 1;
    }
    if (item.indexOf('Actual Length') >= 0) {
      item = `actualLength_${actualLengthCounter}`;
      actualLengthCounter += 1;
    }
    if (item.indexOf('Actual Depth') >= 0) {
      item = `actualDepth_${actualDepthCounter}`;
      actualDepthCounter += 1;
    }
    if (item.indexOf('Dig') >= 0) {
      item = 'digNumber';
    }
    return item;
  });
}

const formatName = (name) => name.toLowerCase().replace(' ', '_');

const createCSV = (sheets, workType) => {
  //looping through all sheets
  let csvStringArray = [];
  for(let x = 0; x < sheets.length; x++) {
      let rows = [];
      let writeStr = "";

      let sheet = sheets[x];
      // only convert sheets with "Crestwood or Glenwood in name"
      if (sheet.name.toLowerCase().indexOf('crestwood') >= 0 || sheet.name.toLowerCase().indexOf('glenwood') >= 0) {
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
  }
  csvStringArray.map((str) => {
    const dir = './outputCSV';
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir);
    }
    fs.writeFile(`${dir}/${str.name}.csv`, str.data, function(err) {
      if(err) {
        return console.log(err);
      }
      console.log(`successfully created CSV file: ${dir}/${str.name}`);
    });
  })
  createJSON(workType)
}

const makeJSON = (file) => new Promise(function(resolve, reject) {
  const dir = './outputJSON';
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir);
  }
  csv().fromFile('./outputCSV/' + file).then((jsonObj) => {
    const outputName = file.split('.')[0] + '.json'
    console.log('outputName', outputName);
    fs.writeFile(`${dir}/${outputName}`, JSON.stringify(jsonObj), (err) => {
      if (err) {
        reject(err);
        throw Error(err)
      }
      resolve('success');
      console.log(`Successfully created ${dir}/${outputName}`);
      cleanJSON(outputName);
    })
  })
});

const createJSON = (workType) => {
  const csvDirectory = './outputCSV/';
  const csvFilePaths = [
    `crestwood_${workType}-complete.csv`,
    `crestwood_${workType}-open.csv`,
    `glenwood_${workType}-complete.csv`,
    `glenwood_${workType}-open.csv`
  ]

  for (let file of csvFilePaths) {
    makeJSON(file);
  }
}

const createTicket = () => {
  return {
    proposedCG: 0,
    ada: 0,
    address: '',
    proposedDimensions: [
      // length, width, height, location
      [0, 0, 0, 'none'],
    ],
    actualDimensions: [
      // length, width, height, location
      [0, 0, 0, 'none'],
    ],
  };
}

const SpreadsheetController = {
  formatCG(cg = '') {
    if (cg && cg.length > 0) {
      const formatted = cg.split('\\s+');
      return parseInt(formatted[0], 10);
    }
    return 0;
  },
  formatDimensions(jsonObj) {
    const formatDim = (val, isDepth = false) => {
      if (typeof val === 'string') {
        if (val === '') {
          return 0;
        }
        if (!isNaN(val)) {
          return parseInt(val);
        }
      }
      return 0;
    };

    const checkIfDepthIsLocation = (depth = '') => {
      if (depth.length > 1 && typeof depth === 'string') {
        return depth;
      }
      return '';
    };

    const {
      proposedLength_0,
      proposedLength_1,
      proposedLength_2,
      proposedWidth_0,
      proposedWidth_1,
      proposedWidth_2,
      proposedDepth_0,
      proposedDepth_1,
      proposedDepth_2,
    } = jsonObj;
    const dim = [
      [
        formatDim(proposedLength_0),
        formatDim(proposedWidth_0),
        formatDim(proposedDepth_0),
        checkIfDepthIsLocation(proposedDepth_0),
      ],
      [
        formatDim(proposedLength_1),
        formatDim(proposedWidth_1),
        formatDim(proposedDepth_1),
        checkIfDepthIsLocation(proposedDepth_1),
      ],
      [
        formatDim(proposedLength_2),
        formatDim(proposedWidth_2),
        formatDim(proposedDepth_2),
        checkIfDepthIsLocation(proposedDepth_2),
      ],
    ];
    const formattedDim = [];
    for (let x = 0; x < dim.length; x += 1) {
      if (dim[x][0] !== 0 || dim[x][1] !== 0 || dim[x][2] !== 0) {
        formattedDim.push(dim[x]);
      }
    }
    return formattedDim;
  },
  formatJSON(json, isConcrete, isCompleted) {
    return json.map((obj) => {
      const ticket = createTicket();

      const proposedCG = SpreadsheetController.formatCG(obj['proposedC&G']);
      const {
        address = '',
        ada = 0,
        area = '',
        town = '',
        digNumber = '',
        releaseDate = '',
        orderNumber = '',
        completionDate = '',
        priority = '',
        issue = '',
        proposedDimensions = [],
      } = obj;
      ticket._p_project = 'nicor';
      ticket.ada = ada === '' ? 0 : ada;
      ticket.address = `${address}, ${town}, IL`;
      ticket.area = area;
      ticket.digNumber = digNumber;
      ticket.isCompleted = isCompleted;
      ticket.completionDate = !completionDate ? new Date() : ExcelDateToJSDate(completionDate);
      ticket.isPriority = priority.length > 0;
      ticket.issueNotes = issue;
      ticket.isWalkedDown = isCompleted;
      ticket.notes = '';
      ticket.orderNumber = orderNumber;
      ticket.pictures = [];
      ticket.proposedCG = proposedCG === '' ? 0 : proposedCG;
      ticket.proposedDimensions = SpreadsheetController.formatDimensions(obj);
      ticket.releaseDate = moment(ExcelDateToJSDate(releaseDate)).toDate();
      ticket.scheduledDate = moment().toDate();
      ticket.status = isCompleted ? 'completed' : 'open';
      ticket.type = isConcrete ? 'Concrete' : 'Asphalt';
      return ticket;
    });
  },
};


const cleanJSON = (file) => {
  console.log('cleanJSON::file', file);
  //EX) crestwood_concrete-complete.json

  const parsedJSON = JSON.parse(fs.readFileSync('./outputJSON/' + file, 'utf8'));
  const isConcrete = file.indexOf('concrete') >= 0;
  const isCompleted = file.indexOf('complete') >= 0;
  const { formatJSON } = SpreadsheetController;
  let cleaned = formatJSON(parsedJSON, isConcrete, isCompleted);
  const dir = './cleanJSON';
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir);
  }
  fs.writeFile(`${dir}/${file}`, JSON.stringify(cleaned), (err) => {
    if (err) {
      throw Error(err)
    }
    console.log(`successfully cleaned JSON: ${dir}/${file}`);
  })
}

rl.question('Which would you like to create JSON for? (concrete or asphalt) ', (answer) => {
  console.log(`You've input: ${answer}`);
  const obj = converter.parse(`./${answer}.xlsx`);
  createCSV(obj, answer);
  rl.close();
});
