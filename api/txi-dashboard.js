// /api/txi-dashboard.js
import { GoogleGenerativeAI } from "@google/generative-ai";

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Use POST with JSON body { question: \"...\" }" });
    }

    const { question } = req.body;
    if (!question) return res.status(400).json({ error: "Missing question" });

    // -----------------------------
    // LOAD ENV VARIABLES
    // -----------------------------
    const SN_URL = process.env.SN_URL;
    const SF_URL = process.env.SF_URL;
    const SP_CHAT_URL = process.env.SP_CHAT_URL;
    const GEMINI_API_KEY = process.env.GEMINI_API_KEY;

    // -----------------------------
    // CALL SERVICE NOW
    // -----------------------------
    let snData = {};
    try {
      const snResp = await fetch(SN_URL);
      snData = await snResp.json();
    } catch (e) {
      snData = { error: e.message };
    }

    // -----------------------------
    // CALL SALESFORCE
    // -----------------------------
    let sfData = {};
    try {
      const sfResp = await fetch(SF_URL);
      sfData = await sfResp.json();
    } catch (e) {
      sfData = { error: e.message };
    }

    // -----------------------------
    // CALL SHAREPOINT
    // -----------------------------
    let spData = {};
    try {
      if (!SP_CHAT_URL) {
        spData = { error: "SP_CHAT_URL not configured" };
      } else {
        const spResp = await fetch(SP_CHAT_URL, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ question }),
        });
        spData = await spResp.json();
      }
    } catch (e) {
      spData = { error: e.message };
    }

    // -----------------------------
    // USE GEMINI 2.5 PRO FOR FINAL SUMMARY
    // -----------------------------
    const genAI = new GoogleGenerativeAI(GEMINI_API_KEY);
    const model = genAI.getGenerativeModel({ model: "gemini-2.5-pro" });

    const aiPrompt = `
You are an executive AI assistant. Summarize risks using:
SALESFORCE DATA: ${JSON.stringify(sfData)}
SERVICENOW DATA: ${JSON.stringify(snData)}
SHAREPOINT DATA: ${JSON.stringify(spData)}

Write a clear leadership-ready answer with:
- Top risks
- Business impact
- Where the signals came from
    `;

    const result = await model.generateContent(aiPrompt);
    const finalText = result.response.text();

    return res.status(200).json({
      question,
      combinedAnswer: finalText,
      sources: {
        salesforce: sfData,
        serviceNow: snData,
        sharePoint: spData,
      },
      generatedAt: new Date().toISOString(),
    });

  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
}
