const express = require('express');
const multer = require('multer');
const path = require('path');
const { processExcel } = require('./processor');

const app = express();
const port = process.env.PORT || 3000;

// Multer
const upload = multer({ storage: multer.memoryStorage() });

// Static files
app.use(express.static(path.join(__dirname, '../public')));
app.use('/output', express.static(path.join(__dirname, '../dist')));

// HOME PAGE
app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "../public/index.html"));
});

// API process
app.post('/api/process', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        const config = JSON.parse(req.body.config);
        const result = await processExcel(req.file.buffer, config);

        res.json(result);
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: error.message });
    }
});

// API analyze
app.post('/api/analyze', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        const { analyzeExcel } = require('./processor');
        const result = await analyzeExcel(req.file.buffer);

        res.json(result);
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: error.message });
    }
});

app.listen(port, () => {
    console.log(`Server running on port ${port}`);
});
