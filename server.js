const fs = require("fs");
const express = require("express");
// const session = require('express-session');
const { google } = require("googleapis");
const { OAuth2Client } = require("google-auth-library");
const path = require("path");
const xlsx = require("xlsx");

const app = express();

// app.use(session({ secret: 'your-secret-key', resave: false, saveUninitialized: true }));

const SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"];
const TOKEN_PATH = "token.json";
const CREDENTIALS_PATH = "credentials.json";

async function authenticate(jsonFile = undefined) {
  if (!jsonFile) {
    const credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH));

    const { client_secret, client_id, redirect_uris } = credentials.web;
    const oAuth2Client = new google.auth.OAuth2(
      client_id,
      client_secret,
      redirect_uris[redirect_uris.length - 1]
    );

    if (fs.existsSync(TOKEN_PATH)) {
      const token = JSON.parse(fs.readFileSync(TOKEN_PATH));
      oAuth2Client.setCredentials(token);
    }

    return oAuth2Client;
  } else {
    const client_id =
      "893178375513-fcqjbusn1lp1q3soctq2bl3s3m03eed7.apps.googleusercontent.com";
    const client_secret = "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";
    const redirect_uris = [
      "https://www.googleapis.com/auth/gmail.readonly",
      "http://localhost:5000/oauth2callback",
    ];
    const oAuth2Client = new google.auth.OAuth2(
      client_id,
      client_secret,
      redirect_uris[redirect_uris.length - 1]
    );

    if (fs.existsSync(TOKEN_PATH)) {
      const token = JSON.parse(fs.readFileSync(TOKEN_PATH));
      oAuth2Client.setCredentials(token);
    }

    return oAuth2Client;
  }
}

async function extractEmails(auth) {
  try {
    const gmail = await getGmailService(auth);
    const res = await gmail.users.labels.list({ userId: "me" });
    const labels = res.data.labels;

    if (!labels.length) {
      return { message: "Nenhuma caixa de correio encontrada." };
    }

    const emails = [];
    let emailIdCounter = 0;

    for (const label of labels) {
      const labelRes = await gmail.users.messages.list({
        userId: "me",
        labelIds: [label.id],
      });
      const messages = labelRes.data.messages || [];

      for (const message of messages) {
        const msg = await gmail.users.messages.get({
          userId: "me",
          id: message.id,
          format: "full",
        });
        const headers = msg.data.payload.headers;

        const emailDetails = {
          ID: emailIdCounter,
          "Caixa de Correio": label.name,
          De: headers.find((header) => header.name === "From")?.value,
          Para: headers.find((header) => header.name === "To")?.value,
          Assunto: headers.find((header) => header.name === "Subject")?.value,
        };

        emails.push(emailDetails);
        emailIdCounter++;
      }
    }

    const ws = xlsx.utils.json_to_sheet(emails);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "Emails");
    xlsx.writeFile(wb, "emails.xlsx");

    return { message: "Extração completa", emails };
  } catch (error) {
    debugger;
    console.log(error);
  }
}

async function saveToken(token) {
  await fs.writeFileSync(TOKEN_PATH, JSON.stringify(token));
}

async function getGmailService(auth) {
  return google.gmail({ version: "v1", auth });
}

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

app.get("/google_login", async (req, res) => {
  const oAuth2Client = await authenticate();
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: "offline",
    scope: SCOPES,
  });
  res.redirect(authUrl);
});

app.get("/oauth2callback", async (req, res) => {
  const oAuth2Client = await authenticate();
  const { code } = req.query;
  const { tokens } = await oAuth2Client.getToken(code);
  oAuth2Client.setCredentials(tokens);
  // await saveToken(tokens);
  await extractEmails(oAuth2Client);
  res.redirect("/download");
  
});

app.get("/download", (req, res) => {
  res.download("emails.xlsx");
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
