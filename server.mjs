import express from 'express';
import cors from 'cors';
import { Buffer } from 'buffer';
import fetch from 'node-fetch';
import dotenv from 'dotenv';

dotenv.config();

const app = express();

app.use(cors({
  origin: [
    'https://localhost:3000',
    'https://storage.googleapis.com', 
    'https://github-jirventures-cube-olap-excel-view-32764122184.us-central1.run.app']
}));

app.use(express.json());

app.post('/api/exportDataSlice', async (req, res) => {
  try {
    const { cubeName, payload, settings } = req.body;
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
