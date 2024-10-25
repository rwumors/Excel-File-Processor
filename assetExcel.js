const ExcelJS = require('exceljs');

// Function to generate the Excel file with styles
function generateExcelFile(assetNumbers) {
    const workbook = new ExcelJS.Workbook();

    // --- Sheet 1: Assets ---
    const assetsSheet = workbook.addWorksheet('Assets', { properties: { tabColor: { argb: '00FF00' } } });
    const assetHexColors = ["FFE699", "FFE699", "D0CECE", "D0CECE", "D0CECE", "FFC000"];

    // Set column headers and widths for Assets sheet
    assetsSheet.columns = [
        { header: 'box #', width: 64 / 7 },
        { header: 'Weight', width: 64 / 7 },
        { header: 'Asset#', width: 64 / 7 },
        { header: 'Serial #', width: 167 / 7 },
        { header: 'Model #', width: 122 / 7 },
        { header: 'Comments / Issues', width: 241 / 7 },
    ];

    // Apply styles to the headers
    assetsSheet.getRow(1).eachCell((cell, colNumber) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: assetHexColors[colNumber - 1] },
        };
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    // Add the asset numbers to the Asset# column
    assetNumbers.forEach(asset => {
        assetsSheet.addRow(["", "", asset, "", "", ""]);
    });

    // --- Sheet 2: RAP ---
    const rapSheet = workbook.addWorksheet('RAP', { properties: { tabColor: { argb: '00B0F0' } } });
    const rapHexColor = "D0CECE";

    // Set column headers and widths for RAP sheet
    rapSheet.columns = [
        { header: 'NET#', width: 75 / 7 },
        { header: 'Serial #', width: 95 / 7 },
        { header: 'MAC Address', width: 95 / 7 },
    ];

    // Apply styles to the headers
    rapSheet.getRow(1).eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: rapHexColor },
        };
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    // --- Sheet 3: DAMAGED Assets ---
    const damagedAssetsSheet = workbook.addWorksheet('DAMAGED Assets', { properties: { tabColor: { argb: 'FF0000' } } });
    const damagedHexColors = ["D0CECE", "D0CECE", "D0CECE", "FFC000"];

    // Set column headers and widths for DAMAGED Assets sheet
    damagedAssetsSheet.columns = [
        { header: 'Asset#', width: 64 / 7 },
        { header: 'Serial #', width: 156 / 7 },
        { header: 'Model#', width: 87 / 7 },
        { header: 'Issue', width: 218 / 7 },
    ];

    // Apply styles to the headers
    damagedAssetsSheet.getRow(1).eachCell((cell, colNumber) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: damagedHexColors[colNumber - 1] },
        };
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    // Return the workbook as a buffer for download
    return workbook.xlsx.writeBuffer();
}

module.exports = { generateExcelFile };
