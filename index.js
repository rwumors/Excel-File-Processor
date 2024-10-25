const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const { processAndUpdateExcelFile } = require('./excelProcessor');
const { generateExcelFile } = require('./assetExcel');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.json());
app.use(express.static('public'));

// Serve the homepage
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/asset', (req, res) => {
    res.sendFile(path.join(__dirname, 'asset.html'));
});

// Endpoint to generate and download the Excel file as a buffer
app.post('/upload', upload.single('serialFile'), async (req, res) => {
    const filePath = req.file.path;
    const ext = path.extname(req.file.originalname).toLowerCase();

    try {
        if (ext !== '.xlsx') {
            return res.status(400).json({ error: 'Please upload an Excel file (.xlsx).' });
        }

        // Process and get updated workbook as a buffer
        const fileBuffer = await processAndUpdateExcelFile(filePath);

        // Set headers to send the buffer as an Excel file
        res.setHeader('Content-Disposition', `attachment; filename=updated_${req.file.originalname}`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        
        // Send the buffer, then delete the file
        res.send(fileBuffer);
        fs.unlinkSync(filePath); // Delete the uploaded file after response is sent

    } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'An error occurred while processing the file.' });
    }
});

// Endpoint to generate a new Excel file from asset numbers
app.post('/assetExcel', async (req, res) => {
    const { assetNumbers } = req.body;

    if (!assetNumbers || assetNumbers.length === 0) {
        return res.status(400).send('No asset numbers provided.');
    }

    try {
        // Generate the Excel file using asset numbers
        const fileBuffer = await generateExcelFile(assetNumbers);

        // Set headers to send the buffer as an Excel file
        res.setHeader('Content-Disposition', 'attachment; filename=Asset_File.xlsx');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        
        // Send the buffer to the client
        res.send(fileBuffer);

    } catch (err) {
        console.error('Error generating Excel file:', err);
        res.status(500).send('Error generating Excel file.');
    }
});

// Start the server
app.listen(3000, () => console.log('Server started on http://localhost:3000'));
