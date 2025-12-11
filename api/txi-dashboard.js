import { GoogleGenerativeAI } from "@google/generative-ai";
import axios from "axios";

export const config = {
  runtime: "nodejs20.x"
};

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Only POST allowed" });
    }

    const question = req.body?.question;
    if (!question) {
      return res.status(400).json({ error: "Missing question" });
    }

    // ---------- ENV VAR VALIDATION ----------
    const SN_URL = process.env.SN_TXI_URL;
    const SN_USER = process.env.SN_USER;
    const SN_PASS = process.env.SN_PASS;

    const SF_URL = process.env.SF_QUERY_URL;
    const SF_TOKEN = process.env.SF_TOKEN;

    const SP_CHAT_URL = process.env.SP_CHAT_URL;

    // wrapper object for output
    let results = {
      salesforce: null,
      serviceNow: null,
      sharePoint: null
    };

    // ---------- SERVICENOW ----------
    try {
      if (!SN_URL || !SN_USER || !SN_PASS) {
        results.serviceNow = {
          source: "ServiceNow",
          error: "ServiceNow env vars not configured (SN_TXI_URL / SN_USERNAME / SN_PASSWORD)."
        };
      } else {
        const sn = await axios.get(SN_URL, {
          auth: { username: SN_USER, password: SN_PASS }
        });
        results.serviceNow = sn.data;
      }
    } catch (err) {
      results.serviceNow = {
        source: "ServiceNow",
        error: err.response?.data || err.message
      };
    }

    // ---------- SALESFORCE ----------
    try {
      if (!SF_URL || !SF_TOKEN) {
        results.salesforce = {
          source: "Salesforce",
          error: "Missing Salesforce env vars (SF_QUERY_URL / SF_TOKEN)"
        };
      } else {
        const sf = await axios.get(SF_URL, {
          headers: { Authorization: `Bearer ${SF_TOKEN}` }
        });
        results.salesforce = sf.data;
      }
    } catch (err) {
      results.salesforce = {
        source: "Salesforce",
        error: err.response?.data || err.message
      };
    }

    // ---------- SHAREPOINT ----------
    try {
      if (!SP_CHAT_URL) {
        results.sharePoint = {
          source: "SharePoint",
          error: "SP_CHAT_URL env var not configured."
        };
      } else {
        const sp = await axios.post(SP_CHAT_URL, { question });
        results.sharePoint = sp.data;
      }
    } catch (err) {
      results.sharePoint = {
        source: "SharePoint",
        error: err.response?.data || err.message
      };
    }

    // ---------- CALL GEMINI ----------
    const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-pro" });

    const prompt = `
You are an executive AI assistant. Summarize operational risks.

Question: "${question}"

ServiceNow Data:
${JSON.stringify(results.serviceNow, null, 2)}

Salesforce Data:
${JSON.stringify(results.salesforce, null, 2)}

SharePoint Data:
${JSON.stringify(results.sharePoint, null, 2)}

Return a clean leadership summary.
`;

    let aiResponse;
    try {
      const r = await model.generateContent(prompt);
      aiResponse = r.response.text();
    } catch (err) {
      aiResponse = null;
    }

    // ---------- FINAL RESPONSE ----------
    return res.status(200).json({
      question,
      combinedAnswer: aiResponse || "Could not generate AI summary.",
      sources: results,
      generatedAt: new Date().toISOString()
    });

  } catch (error) {
    return res.status(500).json({
      error: "Internal Server Error",
      details: error.message
    });
  }
}
