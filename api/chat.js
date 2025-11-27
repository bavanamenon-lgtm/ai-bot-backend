// api/chat.js - Vercel serverless function with CORS enabled

export default async function handler(req, res) {
  // ✅ CORS headers so browser can call this from GitHub Pages / SharePoint
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  // Preflight for browsers
  if (req.method === "OPTIONS") {
    res.status(200).end();
    return;
  }

  try {
    if (req.method !== "POST") {
      res.status(405).json({ error: "Only POST is allowed" });
      return;
    }

    const { question, page_context } = req.body || {};

    if (!question) {
      res.status(400).json({ error: "Missing 'question' in body" });
      return;
    }

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      res
        .status(500)
        .json({ error: "GEMINI_API_KEY is not configured on the server" });
      return;
    }

    const prompt = `
You are an AI assistant embedded inside a SharePoint page.
The current SharePoint context/path is: ${page_context || "unknown"}.

Answer in simple, clear English.
Keep responses short: 3–5 sentences maximum.
If the user asks something unrelated to work, answer politely but briefly.

User question:
${question}
`;

    const body = {
      contents: [
        {
          parts: [{ text: prompt }]
        }
      ]
    };

    const url =
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" +
      apiKey;

    const response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });

    const data = await response.json();

    if (!response.ok) {
      console.error("Gemini error:", response.status, data);
      res.status(500).json({
        error: "Gemini API error",
        status: response.status,
        details: data
      });
      return;
    }

    const answer =
      data?.candidates?.[0]?.content?.parts?.[0]?.text ||
      "Sorry, I couldn’t get a response from the AI.";

    res.status(200).json({ answer });
  } catch (e) {
    console.error("Backend error:", e);
    res.status(500).json({ error: "Server error", details: e.message });
  }
}
