// /api/chat.js

import jsforce from 'jsforce';

const {
  SF_USERNAME,
  SF_PASSWORD,
  SF_TOKEN,
  SF_LOGIN_URL,
  GEMINI_API_KEY
} = process.env;

const loginUrl = SF_LOGIN_URL || 'https://login.salesforce.com';

// ---- Gemini helper ----
async function callGemini(prompt) {
  const url =
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;

  const body = {
    contents: [
      {
        parts: [{ text: prompt }]
      }
    ]
  };

  const resp = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });

  if (!resp.ok) {
    const txt = await resp.text();
    console.error('Gemini error:', resp.status, txt);
    throw new Error('Gemini API error');
  }

  const data = await resp.json();
  const candidate = data.candidates && data.candidates[0];
  const parts = candidate && candidate.content && candidate.content.parts;
  const text = parts && parts[0] && parts[0].text;

  return text || 'I was not able to generate a proper response.';
}

// ---- Simple classifier ----
function classifyQuestion(question) {
  const q = (question || '').toLowerCase();

  if (q.includes('chart') || q.includes('graph') || q.includes('pie')) {
    return 'chart';
  }
  if (q.includes('summary') || q.includes('summarise') || q.includes('summarize')) {
    return 'summary';
  }
  return 'generic';
}

// ---- Main handler ----
export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res
      .status(405)
      .json({ error: 'Use POST with JSON body { "question": "..." }' });
    return;
  }

  try {
    const { question } = req.body || {};
    if (!question || typeof question !== 'string') {
      res.status(400).json({ error: 'Missing "question" in request body.' });
      return;
    }

    // 1) Login to Salesforce
    const conn = new jsforce.Connection({ loginUrl });
    await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);

    // 2) Decide mode + query SF
    const mode = classifyQuestion(question);

    let salesforceRecords = [];
    let chartData = null;

    if (mode === 'chart') {
      // Opportunities grouped by StageName
      const soql =
        "SELECT StageName, COUNT(Id) total " +
        "FROM Opportunity " +
        "WHERE IsClosed = false " +
        "GROUP BY StageName";

      const result = await conn.query(soql);
      salesforceRecords = result.records || [];

      const labels = salesforceRecords.map((r) => r.StageName || 'Unknown');
      const values = salesforceRecords.map((r) =>
        Number(r.total || r.expr0 || 0)
      );

      chartData = {
        title: 'Active opportunities by stage',
        labels,
        values
      };
    } else {
      // Top 5 accounts (used for both summary + generic)
      const soql =
        "SELECT Id, Name, Industry, Rating, Type " +
        "FROM Account " +
        "ORDER BY LastModifiedDate DESC " +
        "LIMIT 5";

      const result = await conn.query(soql);
      salesforceRecords = result.records || [];
    }

    // 3) Build Gemini prompt
    const prompt = `
You are an assistant helping the user understand Salesforce data.

User question:
${question}

Mode: ${mode}

Salesforce data (JSON):
${JSON.stringify(salesforceRecords, null, 2)}

If mode is "chart":
- Briefly explain what the chart would show (1–2 sentences).
- Mention key stages and which stage has the highest and lowest count.

If mode is "summary":
- Summarise in 3–4 sentences.
- Highlight top accounts/opportunities, risks, or anything notable.

Answer in clear, simple English.
`.trim();

    const answer = await callGemini(prompt);

    // 4) Send back to UI
    res.status(200).json({
      answer,
      type: mode,
      chartData,
      salesforceRecords
    });
  } catch (err) {
    console.error('Backend error in /api/chat:', err);
    res.status(500).json({
      error: 'Internal error in /api/chat.',
      details: String(err)
    });
  }
}
