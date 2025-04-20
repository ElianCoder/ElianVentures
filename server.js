const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
require('dotenv').config();
const fetch = require('isomorphic-fetch');
const { Client } = require('@microsoft/microsoft-graph-client');
const rateLimit = require('express-rate-limit');

const app = express();
const PORT = process.env.PORT || 3000;

// CORS: Allow only the frontend domain
app.use(cors({
    origin: ['https://www.elianventures.com'],
    methods: ['POST'],
}));

// Middleware
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// Rate Limiting (Max 10 requests per 15 min)
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 10,
  message: 'Too many requests. Please try again later.',
});
app.use('/request-estimate', limiter);

// Function to get an access token
async function getAccessToken() {
  const url = `https://login.microsoftonline.com/${process.env.OUTLOOK_TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams({
    client_id: process.env.OUTLOOK_CLIENT_ID,
    scope: 'https://graph.microsoft.com/.default',
    client_secret: process.env.OUTLOOK_CLIENT_SECRET,
    grant_type: 'client_credentials',
  });

  try {
    const response = await fetch(url, {
      method: 'POST',
      body: params,
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    });

    if (!response.ok) throw new Error('Failed to fetch access token');

    const data = await response.json();
    return data.access_token;
  } catch (error) {
    console.error('Error in getAccessToken:', error);
    throw error;
  }
}

// Route to handle estimate requests (API Key removed)
app.post('/request-estimate', async (req, res) => {
  const { name, email, message } = req.body;
  if (!name || !email || !message) {
    return res.status(400).send('All fields are required.');
  }

  try {
    const accessToken = await getAccessToken();

    const client = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });

    const emailData = {
      message: {
        subject: `New Estimate Request from ${name}`,
        body: {
          contentType: 'Text',
          content: `You have received a new estimate request:
                   Name: ${name}
                   Email: ${email}
                   Message: ${message}`,
        },
        toRecipients: [
          {
            emailAddress: {
              address: 'IanSantos@elianventures.com',
            },
          },
        ],
      },
      saveToSentItems: false,
    };

    console.log('Sending email...', emailData);
    await client.api(`/users/IanSantos@elianventures.com/sendMail`).post(emailData);

    return res.status(200).send('Estimate request received. We will contact you soon.');
  } catch (error) {
    console.error('Error sending email:', error);
    return res.status(500).send('Something went wrong. Please try again later.');
  }
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});