// excelProcessor.js
const ExcelJS = require('exceljs');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

// Process and update only the "Model#" column in the 'Assets' sheet
async function processAndUpdateExcelFile(filePath) {
    console.log(`Reading file at path: ${filePath}`);
    
    const workbook = new ExcelJS.Workbook();
    const tempFilePath = path.join(__dirname, 'tempWorkbook.xlsx');
    
    try {
        // Load the workbook
        await workbook.xlsx.readFile(filePath);
        console.log("Workbook loaded successfully.");

        const sheet = workbook.getWorksheet('Assets');
        if (!sheet) {
            throw new Error("Assets sheet not found in workbook.");
        }
        console.log(`Processing sheet: ${sheet.name}`);

        // Collect serial numbers from column D where the Model# column (E) is empty
        const serialNumbers = [];
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Skip header row
            const serial = row.getCell('D').value;
            const model = row.getCell('E').value;

            if (serial && !model) {
                serialNumbers.push(serial);
                console.log(`Found serial number without model: ${serial} (Row: ${rowNumber})`);
            }
        });

        const results = await fetchProductNames(serialNumbers);
        const serialNumberToModel = {};
        results.forEach(result => {
            if (result.productName && result.productName !== 'Error: Failed to retrieve data') {
                serialNumberToModel[result.serialNumber] = result.productName;
            }
        });

        // Update the "Model#" column in column E
        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Skip header row
            const serial = row.getCell('D').value;
            const model = row.getCell('E').value;

            if (serial && !model && serialNumberToModel[serial]) {
                row.getCell('E').value = serialNumberToModel[serial];
                console.log(`Updated model for serial ${serial} to ${serialNumberToModel[serial]} at Row ${rowNumber}`);
            }
        });

        // Save workbook to disk temporarily
        await workbook.xlsx.writeFile(tempFilePath);
        console.log("Workbook saved to temporary file successfully.");

        // Re-load workbook to ensure no XML corruption
        const finalWorkbook = new ExcelJS.Workbook();
        await finalWorkbook.xlsx.readFile(tempFilePath);

        // Convert the final, verified workbook to a buffer and clean up
        const buffer = await finalWorkbook.xlsx.writeBuffer();
        fs.unlinkSync(tempFilePath); // Delete temporary file
        console.log("Final workbook processed and written to buffer successfully.");

        return buffer;
    } catch (error) {
        console.error("Error during workbook processing:", error);
        throw new Error("Failed to process workbook.");
    }
}

// Fetch product names from API
async function fetchProductNames(serialNumbers) {
    const url = 'https://pcsupport.lenovo.com/ca/en/api/v4/upsell/redport/getIbaseInfo';
    const headers = { 'Content-Type': 'application/json' };
    const results = [];

    for (const serial of serialNumbers) {
        const body = { serialNumber: serial };
        console.log(`Fetching product name for serial: ${serial}`);
        
        try {
            const response = await axios.post(url, body, { headers });
            let productName = response.data?.data?.machineInfo?.productName || '';

            if (productName.includes('(')) {
                productName = productName.split('(')[0].trim();
            }
            results.push({ serialNumber: serial, productName });
            console.log(`Received product name for ${serial}: ${productName}`);
            
        } catch (error) {
            console.error(`Error fetching product for serial ${serial}:`, error.message);
            results.push({ serialNumber: serial, productName: '' });
        }
    }

    return results;
}

module.exports = { processAndUpdateExcelFile, fetchProductNames };
