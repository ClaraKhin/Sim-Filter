const xlsx = require('xlsx');

const gsmPattern = /^(9|\+?959)(7\d{8}|6\d{8}|5\d{8}|5\d{6}|4\d{7,8}|2\d{6,8}|6\d{6}|8\d{8}|7\d{7}|9(0|1|9)\d{5,6}|2[0-4]\d{5}|5[0-6]\d{5}|8[13-7]\d{5}|4[1379]\d{6}|73\d{6}|91\d{6}|25\d{7}|26[0-5]\d{6}|40[0-4]\d{6}|42\d{7}|45\d{7}|89[6789]\d{6})$/;
const cdmaPattern = /^(9|\+?959)((8\d{6}|6\d{6}|49\d{6})|(3\d{7,8}|73\d{6}|91\d{6})|(47\d{6}))$/;
const callcenterPattern = /^.{4}$/;


//checks:
const checkFormat = (src) => {
  if (!src) {
    return "Invalid";
  }
  const srcStr = src.toString();

  if (callcenterPattern.test(srcStr)) {
    return "HotLine";
  } else if (gsmPattern.test(srcStr)) {
    return "GSM";
  } else if (cdmaPattern.test(srcStr)) {
    return "CDMA";
  } else {
    return "Landline";
  }
};

// Convert Excel’s numeric date to a JS string “YYYY-MM-DD HH:MM:SS”
const convertDate = (excelDate) => {
  if (typeof excelDate !== 'number') return 'Invalid';
  const date = new Date((excelDate - 25569) * 86400 * 1000);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
};


// ——— Parse CLI arguments ———
const args = process.argv;
let inputFile = '';
let outputFile = '';

for (let i = 0; i < args.length; i++) {
  if (args[i] === '-i') {
    inputFile = args[i + 1];
  } else if (args[i] === '-o') {
    outputFile = args[i + 1];
  }
}

if (!inputFile || !outputFile) {
  console.error('Please provide both input (-i) and output (-o) file paths.');
  process.exit(1);
}


// Load the workbook
const workbook = xlsx.readFile(inputFile);

// Analyze all sheets except "Summary"
const sheetNames = workbook.SheetNames.filter(function (sheetName) {
  return sheetName !== 'Summary';
});

// Create objects to store rows for each format type
const formatSheets = {
  HotLine: [],
  GSM: [],
  CDMA: [],
  Landline: []
};

sheetNames.forEach(function (sheetName) {
  const worksheet = workbook.Sheets[sheetName];
  const jsonData = xlsx.utils.sheet_to_json(worksheet, { raw: true });

  // Process each row and categorize by format type
  jsonData.forEach(function (row) {
    // Format the SRC
    const formatType = checkFormat(row.SRC);
    row['SRC_Format'] = formatType;

    // Format row1 if it's a numeric Excel date
    if (typeof row['row1'] === 'number') {
      row['row1'] = formatExcelDate(row['row1']);
    }

    // Add the row to the corresponding format type sheet
    if (formatSheets[formatType]) {
      formatSheets[formatType].push(row);
    }
  });
});

// Create new sheets for each format type and add them to the workbook
Object.keys(formatSheets).forEach(function (formatType) {
  const sheetData = formatSheets[formatType];
  const newWorksheet = xlsx.utils.json_to_sheet(sheetData);
  workbook.Sheets[formatType] = newWorksheet; // Add the new sheet to the workbook
});

// Save the updated workbook
xlsx.writeFile(workbook, outputFile);

console.log(`Processed ${sheetNames.length} sheets and saved to ${outputFile} with separate sheets for each format type.`);