require('dotenv').config();
require('isomorphic-fetch');

const express = require('express')
const app = express()
const port = 3000

app.get('/', (req, res) => {
  res.send('Hello world! Wanna send some mail?');
})

app.get('/send', (req, res) => {
    // Setup variables to connect to MS Graph
    const tenantId = process.env.TENANT_ID
    const clientId = process.env.CLIENT_ID
    const clientSecret = process.env.CLIENT_SECRET
    const senderAddress = process.env.SENDER_ADDRESS
    const recipientAddress = process.env.RECIPIENT_ADDRESS

    // Required for the auth to Graph API
    const scopes = "https://graph.microsoft.com/.default"

    // Setup MS Graph Client and authentication provider
    const { Client } = require("@microsoft/microsoft-graph-client");
    const { TokenCredentialAuthenticationProvider } = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
    const { ClientSecretCredential } = require("@azure/identity");
    const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
        scopes: [scopes]
    });
    const client = Client.initWithMiddleware({
        debugLogging: true,
        authProvider
    });

    // Build email object and send message
    const sendMail = {
        message: {
            subject: 'GitLab Ticket',
            body: {
                contentType: 'Text',
                content: "It's broke, FIX IT!"
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: recipientAddress
                    }
                }
            ]
        },
        saveToSentItems: 'false'
    };

    client.api('/users/' + senderAddress + '/sendMail').post(sendMail);
    res.send('Message sent!');
})

app.listen(port, () => {
  console.log(`Listening on port ${port}`)
})