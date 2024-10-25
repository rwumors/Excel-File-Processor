// excelProcessor.js
const ExcelJS = require('exceljs');
const axios = require('axios');

// Process and update only the "Model#" column in column E, preserving other formatting
async function processAndUpdateExcelFile(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0]; // Use the first sheet

    // Collect serial numbers from column D
    const serialNumbers = [];
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip header row
        const serial = row.getCell('D').value;
        if (serial) serialNumbers.push(serial);
    });

    // Fetch product names for each serial number
    const results = await fetchProductNames(serialNumbers);
    const serialNumberToModel = {};
    results.forEach(result => {
        serialNumberToModel[result.serialNumber] = result.productName;
    });

    // Update only the "Model#" column in column E with the fetched model names
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip header row
        const serial = row.getCell('D').value;
        if (serial && serialNumberToModel[serial]) {
            row.getCell('E').value = serialNumberToModel[serial];
        }
    });

    // Return workbook as a buffer
    return workbook.xlsx.writeBuffer();
}

// Fetch product names from API
async function fetchProductNames(serialNumbers) {
    const url = 'https://pcsupport.lenovo.com/ca/en/api/v4/upsell/redport/getIbaseInfo';
    const headers = { 'Content-Type': 'application/json' };
    const results = [];

    for (const serial of serialNumbers) {
        const body = { serialNumber: serial };
        try {
            const response = await axios.post(url, body, { headers });
            let productName = response.data?.data?.machineInfo?.productName || 'Product name not found';

            if (productName.includes('(')) {
                productName = productName.split('(')[0].trim();
            }
            results.push({ serialNumber: serial, productName });
        } catch (error) {
            console.error(`Error fetching product for serial ${serial}:`, error.message);
            results.push({ serialNumber: serial, productName: 'Error: Failed to retrieve data' });
        }
    }
    return results;
}

module.exports = { processAndUpdateExcelFile, fetchProductNames };
