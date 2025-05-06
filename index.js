const xlsx = require('xlsx');

// Predefined Call Receivers
const cocaCola = 2399111;
const upg = 2399000;
const m150 = 2399215;
const agd = 2399333;

// Regex patterns for number types
const gsmPattern = /^(9|\+?959)(7\d{8}|6\d{8}|5\d{8}|5\d{6}|4\d{7,8}|2\d{6,8}|6\d{6}|8\d{8}|7\d{7}|9(0|1|9)\d{5,6}|2[0-4]\d{5}|5[0-6]\d{5}|8[13-7]\d{5}|4[1379]\d{6}|73\d{6}|91\d{6}|25\d{7}|26[0-5]\d{6}|40[0-4]\d{6}|42\d{7}|45\d{7}|89[6789]\d{6})$/;
const cdmaPattern = /^(9|\+?959)((8\d{6}|6\d{6}|49\d{6})|(3\d{7,8}|73\d{6}|91\d{6})|(47\d{6}))$/;
const callcenterPattern = /^.{4}$/;

// Helper to check SRC format
function checkFormat(src) {
    if (!src) return 'Invalid';
    const srcStr = src.toString();
    if (callcenterPattern.test(srcStr)) return 'HotLine';
    if (gsmPattern.test(srcStr)) return 'GSM';
    if (cdmaPattern.test(srcStr)) return 'CDMA';
    return 'Landline';
}

// Convert Excel numeric date to readable format
function convertDate(excelDate) {
    if (typeof excelDate !== 'number') return 'Invalid';
    const date = new Date((excelDate - 25569) * 86400 * 1000);
    return date.toISOString().replace('T', ' ').slice(0, 19);
}

// CLI Argument Parsing
const args = process.argv;
let inputPath = '';
let outputPath = '';
for (let i = 0; i < args.length; i++) {
    if (args[i] === '-i') inputPath = args[i + 1];
    if (args[i] === '-o') outputPath = args[i + 1];
}

if (!inputPath || !outputPath) {
    console.error('Usage: node index.js -i <input.xlsx> -o <output.xlsx>');
    process.exit(1);
}

// Load workbook
const workbook = xlsx.readFile(inputPath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

const headers = data[0];
const rows = data.slice(1);

// Receiver-specific arrays
const companyData = {
    CocaCola: [],
    UPG: [],
    M150: [],
    AGD: []
};

rows.forEach(row => {
    const callReceiver = row[1];
    if (callReceiver == cocaCola) companyData.CocaCola.push(row);
    else if (callReceiver == upg) companyData.UPG.push(row);
    else if (callReceiver == m150) companyData.M150.push(row);
    else if (callReceiver == agd) companyData.AGD.push(row);
});

// Add Rate and Amount, and determine SRC format
function processData(dataRows) {
    return dataRows.map(row => {
        const rate = 30;
        const billDurationIndex = headers.indexOf('Bill Duration');
        const billDuration = row[billDurationIndex] || 0;
        const amount = Math.floor((billDuration / 60) * rate);
        row.push(rate, amount);

        // Determine SRC_Format (assumes column named "SRC")
        const srcIndex = headers.indexOf('SRC');
        const format = checkFormat(row[srcIndex]);
        row.push(format);

        return row;
    });
}

// Append headers for Rate, Amount, and Format
const extendedHeaders = [...headers, 'Rate', 'Amount', 'SRC_Format'];

// Process each company
Object.entries(companyData).forEach(([companyName, rows]) => {
    const processed = processData(rows);
    const sheet = xlsx.utils.aoa_to_sheet([extendedHeaders, ...processed]);
    workbook.SheetNames.push(companyName);
    workbook.Sheets[companyName] = sheet;
});

// Write output
xlsx.writeFile(workbook, outputPath);
console.log(`File processed and saved to ${outputPath}`);
