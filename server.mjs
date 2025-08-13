import express from 'express';
import cors from 'cors';
import { Buffer } from 'buffer';
import fetch from 'node-fetch';
import dotenv from 'dotenv';

dotenv.config();

// Hardcoded EPM configuration
const EPM_USERNAME = 'itsupport@jirventures.com';
const EPM_PASSWORD = 'Oracle2025@101p!';
const EPM_SERVER_URL = 'https://epmconfluence-test-epmconfluence.epm.us-phoenix-1.ocs.oraclecloud.com';
const EPM_APPLICATION = 'CONFPLAN';

const app = express();

app.use(cors({
  origin: 'https://localhost:3000'
}));

app.use(express.json());

app.post('/api/exportDataSlice', async (req, res) => {
  try {
    const { cubeName, payload } = req.body;
    const url = `${EPM_SERVER_URL}/HyperionPlanning/rest/v3/applications/${EPM_APPLICATION}/plantypes/${cubeName}/exportdataslice`;
    
    // Debug logging
    console.log('Request:', {
      url,
      cubeName,
      payload,
      auth: EPM_USERNAME ? 'present' : 'missing'
    });

    console.log('Request headers:', {
      'Content-Type': 'application/json',
      'Authorization': 'Basic ' + Buffer.from(`${EPM_USERNAME}:${EPM_PASSWORD}`).toString('base64')
    });

    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Basic ' + Buffer.from(`${EPM_USERNAME}:${EPM_PASSWORD}`).toString('base64')
      },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error('Response:', {
        status: response.status,
        text: errorText
      });
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const data = await response.json();
    res.json(data);
  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: error.message });
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`Proxy server running on port ${PORT}`);
});
