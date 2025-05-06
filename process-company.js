const xlsx = require('xlsx');

const cocaCola = 2399111;
const upg = 2399000;
const m150 = 2399215;
const agd = 2399333;

//check the phone numbers of the companies

const checkCompany = (src) => {
    if (!src) {
        return "Invalid";
    }
    const srcStr = src.toString();

    if (srcStr === cocaCola.toString()) {
        return "CocaCola";
    } else if (srcStr === upg.toString()) {
        return "UPG";
    } else if (srcStr === m150.toString()) {
        return "M150";
    } else if (srcStr === agd.toString()) {
        return "AGD";
    } else {
        return "Unknown";
    }
}

// Convert Excel’s numeric date to a JS string “YYYY-MM-DD HH:MM:SS”

const convertDate = (excelDate) => {
    
}