const express = require('express');
const fs = require('fs');
const path = require('path');
const cors = require('cors');

const app = express();
const PORT = 3003; // Use a different port to avoid conflicts with webpack-dev-server

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../..'))); // Serve files from project root

// Endpoint to save prompts
app.post('/save-prompt', (req, res) => {
    try {
        const { filename, content } = req.body;
        
        if (!filename || !content) {
            return res.status(400).json({ 
                error: 'Both filename and content are required' 
            });
        }

        // Ensure the file path is safe and within the project directory
        const safePath = path.resolve(__dirname, '../..', filename);
        const projectRoot = path.resolve(__dirname, '../..');
        
        if (!safePath.startsWith(projectRoot)) {
            return res.status(400).json({ 
                error: 'Invalid file path' 
            });
        }

        // Ensure the directory exists
        const dir = path.dirname(safePath);
        if (!fs.existsSync(dir)) {
            fs.mkdirSync(dir, { recursive: true });
        }

        // Write the file
        fs.writeFileSync(safePath, content, 'utf8');
        
        console.log(`Successfully saved prompt to: ${filename}`);
        res.json({ 
            success: true, 
            message: `Prompt saved to ${filename}`,
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        console.error('Error saving prompt:', error);
        res.status(500).json({ 
            error: 'Failed to save prompt', 
            details: error.message 
        });
    }
});

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ 
        status: 'ok', 
        timestamp: new Date().toISOString(),
        message: 'File server is running'
    });
});

// Start the server
app.listen(PORT, () => {
    console.log(`File server running on http://localhost:${PORT}`);
    console.log(`Health check available at http://localhost:${PORT}/health`);
});

module.exports = app; 