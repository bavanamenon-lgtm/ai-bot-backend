import express from "express";
import fetch from "node-fetch";

const app = express();
app.use(express.json());

// Simple health check
app.get("/", (req, res) => {
  res.send("AI backend is running");
});

// Main chat endpoint
app.post("/api/chat", async (req, res) => {
  try {
    const { question, page_context } = req.body || {};

    if (!question) {
      return res.status(400).json({ error: "Missing 'question' in body" });
    }

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return res
        .status(500)
        .json({ error: "GEMINI_API_KEY is not configured on the server" });
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

    if (!response.ok) {
      const errText = await response.text();
      console.error("Gemini error:", response.status, errText);
      return res
        .status(500)
        .json({ error: "Gemini API error", status: response.status });
    }

    const data = await response.json();
    const answer =
      data?.candidates?.[0]?.content?.parts?.[0]?.text ||
      "Sorry, I couldn’t get a response from the AI.";

    res.json({ answer });
  } catch (e) {
    console.error("Backend error:", e);
    res.status(500).json({ error: "Server error", details: e.message });
  }
});

// Vercel will set PORT; use 3000 locally
const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log("AI backend listening on port", port);
});
