const { ConfidentialClientApplication } = require("@azure/msal-node");
const axios = require("axios");
const dotenv = require("dotenv");

dotenv.config();

const msalConfig = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`, // Your Azure Tenant ID
    clientSecret: process.env.AZURE_CLIENT_SECRET,
  },
};

const cca = new ConfidentialClientApplication(msalConfig);

// Function to get the access token
async function getAccessToken() {
  const tokenRequest = {
    scopes: ["https://graph.microsoft.com/.default"], // MS Graph API scope
  };

  try {
    const response = await cca.acquireTokenByClientCredential(tokenRequest);
    return response?.accessToken;
  } catch (error) {
    console.error("Error fetching access token:", error);
    throw error;
  }
}

// Function to send an email via Microsoft Graph API
async function sendMailGraphAPI(accessToken) {
  const url = `https://graph.microsoft.com/v1.0/users/${process.env.AZURE_USER_ID}/sendMail`;

  const mailData = {
    message: {
      subject: "Test Email from Graph API",
      body: {
        contentType: "HTML",
        content:
          "<p>Hello, this is a test email sent via Microsoft Graph API!</p>",
      },
      toRecipients: [
        {
          emailAddress: {
            address: "rmartins@distology.com",
          },
        },
      ],
    },
    saveToSentItems: true, // Save a copy to the sender's Sent Items
  };

  try {
    const response = await axios.post(url, mailData, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });
    console.log("Email sent successfully:", response.data);
  } catch (error) {
    console.error(
      "Error sending email with Graph API:",
      error.response?.data || error.message,
    );
  }
}

// Main function to get token and send email
async function sendTestEmail() {
  const accessToken = await getAccessToken(); // Fetch access token
  await sendMailGraphAPI(accessToken); // Use Graph API to send email
}

sendTestEmail();
