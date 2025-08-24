import express from 'express';
import cors from 'cors';
import { Buffer } from 'buffer';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import rateLimit from 'express-rate-limit';
import fs from 'fs';
import path from 'path';

dotenv.config();

const app = express();

// Add rate limiting
const limiter = rateLimit({
  windowMs: 1 * 60 * 1000, // 1 minutes
  max: 100, // Limit each IP to 100 requests per windowMs
  standardHeaders: true, // Return rate limit info in the `RateLimit-*` headers
  legacyHeaders: false, // Disable the `X-RateLimit-*` headers
  message: 'Too many requests from this IP, please try again after 1 minute',
});

// Apply the rate limiting middleware to API calls only
app.use('/api', limiter);

// Configure CORS to allow requests from Excel add-in origins
const corsOptions = {
  origin: [
    'https://storage.googleapis.com',
    'https://excel.officeapps.live.com',
    'https://excel.office.com',
    'https://outlook.office.com',
    'https://outlook.office365.com',
    'https://outlook.live.com',
    'null', // For local development/file:// origins
    /^https:\/\/.*\.officeapps\.live\.com$/,
    /^https:\/\/.*\.office\.com$/,
    /^https?:\/\/localhost:\d+$/ // Match any localhost port
  ],
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: [
    'Content-Type',
    'Authorization',
    'X-Requested-With',
    'Accept',
    'Origin'
  ],
  credentials: true,
  optionsSuccessStatus: 200 // Some legacy browsers (IE11, various SmartTVs) choke on 204
};

app.use(cors(corsOptions));

app.use(express.json());

// Centralized error handler middleware
const errorHandler = (err, req, res, next) => {
  console.error('An unexpected error occurred:', err);
  if (res.headersSent) {
    return next(err);
  }
  res.status(500).json({ error: 'An internal server error occurred.' });
};

app.post('/api/exportDataSlice', async (req, res, next) => {
  try {
    const { cubeName, payload, settings } = req.body;

    if (!cubeName || !payload || !settings) {
      return res.status(400).json({ message: 'Missing required parameters: cubeName, payload, or settings.' });
    }

    let url;

    if (settings.connectionType === 'hyperion') {
      url = `${settings.serverUrl}/HyperionPlanning/rest/v3/applications/${settings.application}/plantypes/${cubeName}/exportdataslice`;
    } else {
      url = settings.serverUrl.replace('{cube_name}', cubeName);
    }
    
    // Debug logging
    console.log('Request:', {
      url,
      cubeName,
      payload,
      auth: settings.username ? 'present' : 'missing'
    });

    const fetchOptions = {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
    };

    if (settings.connectionType === 'hyperion') {
        fetchOptions.headers['Authorization'] = 'Basic ' + Buffer.from(`${settings.username}:${settings.password}`).toString('base64');
    }

    console.log('Request headers:', fetchOptions.headers);

    const response = await fetch(url, fetchOptions);
    const responseBody = await response.text();

    if (!response.ok) {
      console.error('Response:', {
        status: response.status,
        text: responseBody
      });
      try {
        const errorJson = JSON.parse(responseBody);
        res.status(response.status).json(errorJson);
      } catch (e) {
        res.status(response.status).json({ message: responseBody });
      }
      return;
    }

    try {
      const data = JSON.parse(responseBody);
      res.json(data);
    } catch (error) {
      console.error('Error parsing JSON:', error);
      res.status(500).json({ error: 'Invalid JSON response from server' });
    }
  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/saveReport', (req, res) => {
    const report = req.body;
    const reportDir = path.join(__dirname, 'reports');

    if (!fs.existsSync(reportDir)) {
        fs.mkdirSync(reportDir);
    }

    const fileName = `${report.name.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_v${report.version}.json`;
    const filePath = path.join(reportDir, fileName);

    fs.writeFile(filePath, JSON.stringify(report, null, 2), (err) => {
        if (err) {
            console.error('Error saving report:', err);
            return res.status(500).json({ message: 'Failed to save report' });
        }
        console.log(`Report saved to ${filePath}`);
        res.status(200).json({ message: 'Report saved successfully' });
    });
});

// Use the centralized error handler
app.use(errorHandler);

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`Proxy server running on port ${PORT}`);
});
