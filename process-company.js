const xlsx = require('xlsx');

// Predefined numbers
const cocaCola = 2399111;
const upg = 2399000;
const m150 = 2399215;
const agd = 2399333;

// Function to process and separate the Excel file into different sheets
function processExcelFile(inputPath, outputPath) {
    // Read the Excel file
    const workbook = xlsx.readFile(inputPath);
    const sheetName = workbook.SheetNames[0]; // Assuming the first sheet is the target
    const worksheet = workbook.Sheets[sheetName];

    // Convert the sheet to JSON
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    // Extract headers and rows
    const headers = data[0];
    const rows = data.slice(1);

    // Separate data based on "Call Receiver"
    const cocaColaData = [];
    const upgData = [];
    const m150Data = [];
    const agdData = [];

    rows.forEach(function (row) {
        const callReceiver = row[1]; // Assuming "Call Receiver" is in column B (index 1)

        if (callReceiver == cocaCola) {
            cocaColaData.push(row);
        } else if (callReceiver == upg) {
            upgData.push(row);
        } else if (callReceiver == m150) {
            m150Data.push(row);
        } else if (callReceiver == agd) {
            agdData.push(row);
        }
    });

    // Function to process data and add "Rate" and "Amount"
    function processData(data) {
        return data.map(function (row) {
            // Add "Rate" column with value 30
            const rate = 30;
            row.push(rate);

            // Calculate "Amount" based on "Bill Duration" and "Rate"
            const billDurationIndex = headers.indexOf('Bill Duration');
            const billDuration = row[billDurationIndex] || 0; // Default to 0 if missing
            const amount = Math.floor((billDuration / 60) * rate);
            row.push(amount);

            return row;
        });
    }

    // Process each dataset
    const processedCocaColaData = processData(cocaColaData);
    const processedUpgData = processData(upgData);
    const processedM150Data = processData(m150Data);
    const processedAgdData = processData(agdData);

    // Create new sheets for each dataset
    const cocaColaSheet = xlsx.utils.aoa_to_sheet([headers, ...processedCocaColaData]);
    const upgSheet = xlsx.utils.aoa_to_sheet([headers, ...processedUpgData]);
    const m150Sheet = xlsx.utils.aoa_to_sheet([headers, ...processedM150Data]);
    const agdSheet = xlsx.utils.aoa_to_sheet([headers, ...processedAgdData]);

    // Add the new sheets to the workbook
    workbook.SheetNames.push('CocaCola', 'UPG', 'M150', 'AGD');
    workbook.Sheets['CocaCola'] = cocaColaSheet;
    workbook.Sheets['UPG'] = upgSheet;
    workbook.Sheets['M150'] = m150Sheet;
    workbook.Sheets['AGD'] = agdSheet;

    // Write the updated workbook to the output file
    xlsx.writeFile(workbook, outputPath);

    console.log('Excel file has been updated and saved to ' + outputPath);
}

// Parse command-line arguments
const args = process.argv.slice(2);
let inputPath = '';
let outputPath = '';

// Parse input and output file paths
for (let i = 0; i < args.length; i++) {
    if (args[i] === '-i') {
        inputPath = args[i + 1];
    } else if (args[i] === '-o') {
        outputPath = args[i + 1];
    }
}

// Check for required arguments
if (!inputPath || !outputPath) {
    console.error('Usage: node your_script.js -i <input-file> -o <output-file>');
    process.exit(1);
}

// Process the Excel file
processExcelFile(inputPath, outputPath);