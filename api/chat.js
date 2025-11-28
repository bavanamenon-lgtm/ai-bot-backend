// api/chat.js
import jsforce from "jsforce";
import fetch from "node-fetch";

export default async function handler(req, res) {
  // --- CORS for Salesforce web tab / LWC host ---
  res.setHeader("Access-Control-Allow-Origin", "*"); // for POC; later restrict to your SF domain
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  // Handle preflight
  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  // Allow only POST
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Only POST is allowed" });
  }

  try {
    const { question } = req.body || {};
    if (!question) {
      return res
        .status(400)
        .json({ error: "Missing 'question' in request body" });
    }

    // --- 1. Connect to Salesforce using username + password + token ---
    const username = process.env.SF_USERNAME;
    const password = process.env.SF_PASSWORD;
    const token = process.env.SF_TOKEN || "";
    const loginUrl = process.env.SF_LOGIN_URL || "https://login.salesforce.com";

    if (!username || !password) {
      throw new Error("SF_USERNAME or SF_PASSWORD is not configured");
    }

    const conn = new jsforce.Connection({ loginUrl });

    // password + token (classic basic auth pattern)
    await conn.login(username, password + token);

    // Simple sample query – top 5 Accounts
    const result = await conn.query(
      "SELECT Id, Name, BillingCountry, Industry FROM Account LIMIT 5"
    );

    const sfText = JSON.stringify(result.records, null, 2);

    // --- 2. Ask Gemini to summarise / answer based on Salesforce data ---

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      throw new Error("GEMINI_API_KEY is not configured");
    }

    const prompt = `
You are an AI assistant connected to Salesforce CRM.

User question:
${question}

Salesforce query result (top 5 Accounts):
${sfText}

Answer the user in simple English, 3–4 sentences maximum.
If the user asks to list accounts, give a short bullet list with Name and Country.
`;

    const gRes = await fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" +
        apiKey,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
        }),
      }
    );

    if (!gRes.ok) {
      const bodyText = await gRes.text();
      throw new Error(
        `Gemini API error: ${gRes.status} ${gRes.statusText} – ${bodyText}`
      );
    }

    const data = await gRes.json();
    const answer =
      data?.candidates?.[0]?.content?.parts?.[0]?.text ||
      "I couldn't generate an answer.";

    // --- 3. Return answer + raw Salesforce rows (handy for debugging) ---
    return res.status(200).json({
      answer,
      salesforceRecords: result.records,
    });
  } catch (err) {
    console.error("API /api/chat error:", err);
    return res.status(500).json({
      error: err.message || "Server error",
    });
  }
}
