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

// Endpoint to generate and download the Excel file as buffer
app.post('/upload', upload.single('serialFile'), async (req, res) => {
    const filePath = req.file.path;
    const ext = path.extname(req.file.originalname).toLowerCase();

    try {
        if (ext !== '.xlsx') {
            return res.status(400).json({ error: 'Please upload an Excel file (.xlsx).' });
        }

        // Process and get updated workbook as a buffer
        const fileBuffer = await processAndUpdateExcelFile(filePath);

        // Send the file as a downloadable response
        res.setHeader('Content-Disposition', 'attachment; filename=updated_' + req.file.originalname);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(fileBuffer);

        // Clean up the uploaded file
        fs.unlinkSync(filePath);

    } catch (err) {
        console.error(err);
        res.status(500).json({ error: 'An error occurred while processing the file.' });
    }
});

// Start the server
app.listen(3000, () => console.log('Server started on http://localhost:3000'));
