// api/chat.js

import jsforce from "jsforce";

// ---------- Salesforce connection helper ----------
async function connectSF() {
  const username = process.env.SF_USERNAME;
  const password = process.env.SF_PASSWORD;
  const token = process.env.SF_TOKEN;

  if (!username || !password || !token) {
    throw new Error(
      "Salesforce credentials are missing. Please set SF_USERNAME, SF_PASSWORD and SF_TOKEN in Vercel."
    );
  }

  const conn = new jsforce.Connection({
    loginUrl: "https://login.salesforce.com",
  });

  await conn.login(username, password + token);
  return conn;
}

// ---------- Main API handler ----------
export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Only POST is allowed" });
    return;
  }

  const { question } = req.body || {};

  if (!question) {
    res.status(400).json({ error: "Missing 'question' in request body" });
    return;
  }

  try {
    // 1) Connect to Salesforce
    const conn = await connectSF();

    // 2) Simple demo query – you can change this to any object / SOQL
    const sfResult = await conn.query(`
      SELECT Id, Name, Industry, BillingCity
      FROM Account
      LIMIT 5
    `);

    const sfText = JSON.stringify(sfResult.records, null, 2);

    // 3) Build prompt for Gemini
    const prompt = `
You are an AI assistant connected to Salesforce CRM.

User question:
${question}

Salesforce data (JSON):
${sfText}

Using ONLY the Salesforce data above, answer the user's question
in simple English, in 3–4 sentences. If the answer is not in the data,
say briefly that you cannot find it in Salesforce.
    `;

    // 4) Call Gemini
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      throw new Error(
        "GEMINI_API_KEY is not configured. Please set it in Vercel Environment Variables."
      );
    }

    const geminiUrl =
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" +
      apiKey;

    const geminiResponse = await fetch(geminiUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
      }),
    });

    if (!geminiResponse.ok) {
      const errText = await geminiResponse.text();
      console.error("Gemini error:", geminiResponse.status, errText);
      res.status(500).json({
        error: "Gemini API error",
        status: geminiResponse.status,
        details: errText,
      });
      return;
    }

    const geminiData = await geminiResponse.json();
    const answer =
      geminiData?.candidates?.[0]?.content?.parts?.[0]?.text ||
      "Sorry, I couldn't generate an answer from Salesforce data.";

    // 5) Send answer back to frontend
    res.status(200).json({
      answer,
      sfRecords: sfResult.records, // optional: for debugging or display
    });
  } catch (err) {
    console.error("API /api/chat error:", err);
    res.status(500).json({
      error: "Server error",
      details: err.message,
    });
  }
}
