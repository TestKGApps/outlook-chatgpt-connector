const express = require('express');
const axios = require('axios');
const cors = require('cors');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());

// Get access token for Microsoft Graph API
async function getAccessToken(refreshToken) {
    const response = await axios.post(`https://login.microsoftonline.com/d41ace3b-e558-4d75-8a5a-ef52fd57be2c/oauth2/v2.0/token`, {
          client_id: process.env.CLIENT_ID,
          client_secret: process.env.CLIENT_SECRET,
          grant_type: 'refresh_token',
          refresh_token: refreshToken,
          scope: 'https://graph.microsoft.com/.default'
    });
    return response.data.access_token;
}

// Search emails endpoint
app.post('/api/search-emails', async (req, res) => {
    try {
          const { query, refreshToken } = req.body;
          const token = await getAccessToken(refreshToken);
          const response = await axios.get(
                  `https://graph.microsoft.com/v1.0/me/messages?$search="${query}"&$top=20`,
            { headers: { Authorization: `Bearer ${token}` } }
                );
          res.json(response.data);
    } catch (error) {
          res.status(500).json({ error: error.message });
    }
});

// Get folder statistics
app.post('/api/folder-stats', async (req, res) => {
    try {
          const { refreshToken } = req.body;
          const token = await getAccessToken(refreshToken);
          const response = await axios.get(
                  `https://graph.microsoft.com/v1.0/me/mailFolders`,
            { headers: { Authorization: `Bearer ${token}` } }
                );
          res.json(response.data);
    } catch (error) {
          res.status(500).json({ error: error.message });
    }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
