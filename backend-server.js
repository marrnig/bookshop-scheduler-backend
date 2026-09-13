const express = require('express');
const axios = require('axios');
require('dotenv').config();

const app = express();
app.use(express.json());

const cors = require('cors');
app.use(cors());

// Configuration
const CONFIG = {
  CLIENT_ID: process.env.CLIENT_ID,
  CLIENT_SECRET: process.env.CLIENT_SECRET,
  TENANT_ID: process.env.TENANT_ID,
  SHAREPOINT_SITE_ID: process.env.SHAREPOINT_SITE_ID,
  SHIFTS_LIST_ID: process.env.SHIFTS_LIST_ID || '8ecdea9d-9f63-40fd-9503-26319aaca57d',
  AVAILABILITY_LIST_ID: process.env.AVAILABILITY_LIST_ID || '73cdb605-1ce6-4647-ae69-54c926fd0fe2',
  SCHEDULE_LIST_ID: process.env.SCHEDULE_LIST_ID || '6a4372ec-4e6a-46f3-8499-7582b5cab075',
};

// Validate config
const requiredConfig = ['CLIENT_ID', 'CLIENT_SECRET', 'TENANT_ID', 'SHAREPOINT_SITE_ID'];
for (const key of requiredConfig) {
  if (!CONFIG[key]) {
    console.error(`Missing required environment variable: ${key}`);
    process.exit(1);
  }
}

let cachedToken = null;
let tokenExpiry = null;

// Get OAuth token from Azure AD
async function getAccessToken() {
  // Return cached token if still valid
  if (cachedToken && tokenExpiry > Date.now()) {
    return cachedToken;
  }

  try {
    const response = await axios.post(
      `https://login.microsoftonline.com/${CONFIG.TENANT_ID}/oauth2/v2.0/token`,
      new URLSearchParams({
        client_id: CONFIG.CLIENT_ID,
        client_secret: CONFIG.CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials',
      })
    );

    cachedToken = response.data.access_token;
    tokenExpiry = Date.now() + (response.data.expires_in * 1000) - 60000; // Refresh 1 min before expiry
    return cachedToken;
  } catch (error) {
    console.error('Error getting access token:', error.response?.data || error.message);
    throw new Error('Failed to authenticate with SharePoint');
  }
}

// Get SharePoint items (with optional filter)
async function getSharePointItems(listId, filter = null) {
  const token = await getAccessToken();
  const graphUrl = `https://graph.microsoft.com/v1.0/sites/${CONFIG.SHAREPOINT_SITE_ID}/lists/${listId}/items?$expand=fields`;

  try {
    const response = await axios.get(graphUrl, {
      headers: { Authorization: `Bearer ${token}` },
      params: filter ? { $filter: filter } : {},
    });
    return response.data.value.map(item => item.fields);
  } catch (error) {
    console.error('Error reading SharePoint:', error.response?.data || error.message);
    throw error;
  }
}

// Write to SharePoint list
async function writeToSharePoint(listId, fields) {
  const token = await getAccessToken();
  const graphUrl = `https://graph.microsoft.com/v1.0/sites/${CONFIG.SHAREPOINT_SITE_ID}/lists/${listId}/items`;

  try {
    const response = await axios.post(
      graphUrl,
      { fields },
      { headers: { Authorization: `Bearer ${token}` } }
    );
    return response.data;
  } catch (error) {
    console.error('Error writing to SharePoint:', error.response?.data || error.message);
    throw error;
  }
}

// Routes

// Health check
app.get('/health', (req, res) => {
  res.json({ status: 'ok' });
});

// Get upcoming shifts
app.get('/api/shifts', async (req, res) => {
  try {
    const shifts = await getSharePointItems(CONFIG.SHIFTS_LIST_ID);

    // Sort by date
    shifts.sort((a, b) => new Date(a.ShiftDate) - new Date(b.ShiftDate));

    res.json(shifts);
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch shifts' });
  }
});

// Get list of colleagues (for pairing dropdown)
app.get('/api/colleagues', async (req, res) => {
  try {
    const availabilities = await getSharePointItems(CONFIG.AVAILABILITY_LIST_ID);

    // Get unique people with their emails
    const colleagues = {};
    availabilities.forEach(avail => {
      if (avail.PersonEmail && !colleagues[avail.PersonEmail]) {
        colleagues[avail.PersonEmail] = {
          email: avail.PersonEmail,
          name: avail.PersonEmail.split('@')[0], // Use part before @ as name
        };
      }
    });

    res.json(Object.values(colleagues));
  } catch (error) {
    res.status(500).json({ error: 'Failed to fetch colleagues' });
  }
});

// Save availability
app.post('/api/availability', async (req, res) => {
  const { email, availability } = req.body;

  // Validate email
  if (!email || !email.includes('@')) {
    return res.status(400).json({ error: 'Invalid email address' });
  }

  if (!availability || typeof availability !== 'object') {
    return res.status(400).json({ error: 'Invalid availability data' });
  }

  try {
    // For each slot, create or update availability record
    const slots = Object.entries(availability);
    const results = [];

    for (const [slotId, slotData] of slots) {
      if (!slotData.available) continue; // Skip unavailable slots

      const parts = slotId.split('-');
      const timeSlot = parts[parts.length - 1]; // Last part (am/pm)
      const dateStr = parts.slice(0, -1).join('-'); // Everything before (date)

      // Find matching shift
      const shifts = await getSharePointItems(CONFIG.SHIFTS_LIST_ID);
            
      const shift = shifts.find(s =>
        s.ShiftDate.split('T')[0] === dateStr &&
        s.TimeSlot?.toLowerCase() === timeSlot?.toLowerCase()
      );

      if (!shift) {
        console.warn(`No shift found for ${slotId}`);
        continue;
      }

      // Create availability record
      const title = `${email} - ${shift.TimeSlot} ${shift.ShiftDate.split('T')[0]}`;
      const fields = {
        Title: title,
        PersonEmail: email,
        ShiftID: shift.id, // Link to shift
        Available: true,
        PairWithEmail: slotData.pairWith || null,
        IncludesPlus1: slotData.hasPlus1 || false,
        Plus1Name: slotData.plus1Name || null,
      };

      fields.ShiftID = shift.id;
      const result = await writeToSharePoint(CONFIG.AVAILABILITY_LIST_ID, fields);
      results.push(result);
    }

    res.json({
      success: true,
      message: `Saved ${results.length} availability slot(s)`,
      count: results.length
    });
  } catch (error) {
    console.error('Error saving availability:', error);
    res.status(500).json({ error: 'Failed to save availability' });
  }
});

// Serve the form (you can also host this separately)
app.get('/', (req, res) => {
  res.send('API is running. Use POST /api/availability to submit shifts.');
});

// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Bookshop Shift Scheduler API listening on port ${PORT}`);
  console.log(`Environment: ${process.env.NODE_ENV || 'development'}`);
});
