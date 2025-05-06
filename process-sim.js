const xlsx = require('xlsx');

// Define GSM and CDMA regex patterns
const gsmPattern = /^(9|\+?959)(7\d{8}|6\d{8}|5\d{8}|5\d{6}|4\d{7,8}|2\d{6,8}|6\d{6}|8\d{8}|7\d{7}|9(0|1|9)\d{5,6}|2[0-4]\d{5}|5[0-6]\d{5}|8[13-7]\d{5}|4[1379]\d{6}|73\d{6}|91\d{6}|25\d{7}|26[0-5]\d{6}|40[0-4]\d{6}|42\d{7}|45\d{7}|89[6789]\d{6})$/;
const cdmaPattern = /^(9|\+?959)((8\d{6}|6\d{6}|49\d{6})|(3\d{7,8}|73\d{6}|91\d{6})|(47\d{6}))$/;
const callcenterPattern = /^.{4}$/;

// Function to check the format of SRC numbers
function checkFormat(src) {
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
}

// Function to format date values
function formatExcelDate(excelDate) {
  if (typeof excelDate !== 'number') return 'Invalid';
  const date = new Date((excelDate - 25569) * 86400 * 1000); // Excel's epoch starts at 25569 days after 1900
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

// Parse command-line arguments
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

// Load Excel file
const workbook = xlsx.readFile(inputFile);

// Analyze all sheets except "Summary"
const sheetNames = workbook.SheetNames.filter(function (sheetName) {
  return sheetName !== 'Summary';
});

sheetNames.forEach(function (sheetName) {
  const worksheet = workbook.Sheets[sheetName];
  const jsonData = xlsx.utils.sheet_to_json(worksheet, { raw: true }); // Keep raw values for further processing

  // Add a new 'SRC_Format' column and format the date values
  jsonData.forEach(function (row) {
    // Format the SRC
    row['SRC_Format'] = checkFormat(row.SRC);

    // Format row1 if it's a numeric Excel date
    if (typeof row['row1'] === 'number') {
      row['row1'] = formatExcelDate(row['row1']);
    }
  });

  // Convert the JSON data back to a worksheet
  const updatedWorksheet = xlsx.utils.json_to_sheet(jsonData);

  // Replace the original sheet with the updated one
  workbook.Sheets[sheetName] = updatedWorksheet;
});

// Save the updated workbook
xlsx.writeFile(workbook, outputFile);

console.log('Analysis complete and saved to ' + outputFile);