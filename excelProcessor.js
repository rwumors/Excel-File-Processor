// excelProcessor.js
const ExcelJS = require('exceljs');
const axios = require('axios');

// Process and update only the "Model#" column in column E, preserving other formatting
async function processAndUpdateExcelFile(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0]; // Use the first sheet

    // Collect serial numbers from column D where the Model# column (E) is empty
    const serialNumbers = [];
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip header row
        const serial = row.getCell('D').value;
        const model = row.getCell('E').value;

        // Only collect serials where Model# is empty
        if (serial && !model) serialNumbers.push(serial);
    });

    // Fetch product names for each serial number
    const results = await fetchProductNames(serialNumbers);
    const serialNumberToModel = {};
    results.forEach(result => {
        // Only add to map if a valid product name was found
        if (result.productName && result.productName !== 'Error: Failed to retrieve data') {
            serialNumberToModel[result.serialNumber] = result.productName;
        }
    });

    // Update only the "Model#" column in column E with the fetched model names
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip header row
        const serial = row.getCell('D').value;
        const model = row.getCell('E').value;

        // Update Model# only if it's currently empty and we have a model name for the serial
        if (serial && !model && serialNumberToModel[serial]) {
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
            let productName = response.data?.data?.machineInfo?.productName || '';

            // Extract the first part of the product name before the first '(' character
            if (productName.includes('(')) {
                productName = productName.split('(')[0].trim();
            }
            results.push({ serialNumber: serial, productName });
        } catch (error) {
            console.error(`Error fetching product for serial ${serial}:`, error.message);
            results.push({ serialNumber: serial, productName: '' });
        }
    }
    return results;
}

module.exports = { processAndUpdateExcelFile, fetchProductNames };
