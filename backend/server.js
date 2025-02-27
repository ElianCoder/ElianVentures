const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
require('dotenv/lib/main.js').config(); // Load environment variables
const fetch = require('isomorphic-fetch');
const { Client } = require('@microsoft/microsoft-graph-client');

const app = express();
const PORT = process.env.PORT || 3000;

// Enable CORS for all routes
app.use(cors());

// Middleware to parse JSON and form data
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// Function to get an access token
async function getAccessToken() {
  const url = `https://login.microsoftonline.com/${process.env.OUTLOOK_TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams({
    client_id: process.env.OUTLOOK_CLIENT_ID,
    scope: 'https://graph.microsoft.com/.default',
    client_secret: process.env.OUTLOOK_CLIENT_SECRET,
    grant_type: 'client_credentials',
  });

  console.log('Fetching access token...'); // Debug log

  try {
    const response = await fetch(url, {
      method: 'POST',
      body: params,
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error('Error fetching access token:', errorText);
      throw new Error(`Failed to fetch access token: ${response.statusText}`);
    }

    const data = await response.json();
    console.log('Access token retrieved:', data.access_token ? 'Success' : 'Failed');
    console.log('Access token retrieved:', data.access_token);

    return data.access_token;
  } catch (error) {
    console.error('Error in getAccessToken:', error);
    throw error;
  }
}

// Route to handle estimate requests
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
              address: 'IanSantos@ElianVenturesLLC.onmicrosoft.com', // Your business email
            },
          },
        ],
      },
      saveToSentItems: false,
    };
    console.log('Sending email...',emailData); // Debug log
    // Use '/users/{your-email}/sendMail' instead of '/me/sendMail'
    await client.api(`/users/IanSantos@ElianVenturesLLC.onmicrosoft.com/sendMail`).post(emailData);

    console.log('Email sent successfully');
    return res.status(200).send('Estimate request received. We will contact you soon.');
  } catch (error) {
    console.error('Error sending email:', error);
    return res.status(500).send(`Error sending email: ${error.message}`);
  }
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});