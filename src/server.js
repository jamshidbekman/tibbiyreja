const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { processExcel } = require('./processor');

const app = express();
const port = 3000;

// Configure Multer for memory storage (process in memory)
const upload = multer({ storage: multer.memoryStorage() });

// Serve static files
app.use(express.static(path.join(__dirname, '../public')));
app.use('/output', express.static(path.join(__dirname, '../dist')));

// API Endpoint
app.post('/api/process', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        // Parse config from body (it comes as stringified JSON)
        const config = JSON.parse(req.body.config);

        console.log('Processing file:', req.file.originalname);
        console.log('Config:', JSON.stringify(config, null, 2));

        const result = await processExcel(req.file.buffer, config);

        res.json(result);
    } catch (error) {
        console.error('Error processing:', error);
        res.status(500).json({ error: error.message });
    }
});

// API Endpoint for Analysis
app.post('/api/analyze', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        // We reuse the analyze logic from processor
        const { analyzeExcel } = require('./processor');
        const result = await analyzeExcel(req.file.buffer);

        res.json(result);
    } catch (error) {
        console.error('Error analyzing:', error);
        res.status(500).json({ error: error.message });
    }
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
