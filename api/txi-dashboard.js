import { GoogleGenerativeAI } from "@google/generative-ai";
import axios from "axios";

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Only POST allowed" });
    }

    const { question } = req.body;
    if (!question) {
      return res.status(400).json({ error: "Missing question" });
    }

    // -----------------------------
    // ENV VARS
    // -----------------------------
    const {
      GEMINI_API_KEY,
      SN_URL,
      SN_USERNAME,
      SN_PASSWORD,
      SF_USERNAME,
      SF_PASSWORD,
      SF_TOKEN,
      SF_LOGIN_URL,
      SP_CHAT_URL
    } = process.env;

    // -----------------------------
    // SERVICE NOW FETCH
    // -----------------------------
    let serviceNowData;
    try {
      const sn = await axios.get(`${SN_URL}/api/dtp/schindler_txi/incident_summary`, {
        auth: { username: SN_USERNAME, password: SN_PASSWORD }
      });
      serviceNowData = sn.data;
    } catch (e) {
      serviceNowData = { error: e.message, source: "ServiceNow" };
    }

    // -----------------------------
    // SALESFORCE FETCH
    // -----------------------------
    let salesforceData;
    try {
      const loginRes = await axios.post(`${SF_LOGIN_URL}/services/Soap/u/58.0`, {
        username: SF_USERNAME,
        password: SF_PASSWORD + SF_TOKEN
      });

      const sessionId =
        loginRes.data.match(/<sessionId>(.*?)<\/sessionId>/)?.[1] || null;
      const serverUrl =
        loginRes.data.match(/<serverUrl>(.*?)<\/serverUrl>/)?.[1].replace(
          "/services/Soap/u/58.0",
          ""
        );

      const sf = axios.create({
        baseURL: serverUrl,
        headers: { Authorization: `Bearer ${sessionId}` }
      });

      const atRiskQuery = `
        SELECT Id, Name, Amount, StageName, Probability, CloseDate
        FROM Opportunity
        WHERE Is_At_Risk__c = true
      `;

      const accQuery = `
        SELECT Id, Name, Industry, Rating
        FROM Account
        WHERE Name = 'EBC HQ'
        LIMIT 1
      `;

      const [accRes, riskRes] = await Promise.all([
        sf.get(`/services/data/v58.0/query?q=${encodeURIComponent(accQuery)}`),
        sf.get(`/services/data/v58.0/query?q=${encodeURIComponent(atRiskQuery)}`)
      ]);

      salesforceData = {
        source: "Salesforce",
        ebcAccount: accRes.data.records?.[0] || null,
        atRiskOpportunities: riskRes.data.records || [],
        atRiskSummary: {
          opportunityCount: riskRes.data.records.length,
          totalAmount: riskRes.data.records.reduce((t, r) => t + (r.Amount || 0), 0)
        }
      };
    } catch (e) {
      salesforceData = { source: "Salesforce", error: e.message };
    }

    // -----------------------------
    // SHAREPOINT FETCH
    // -----------------------------
    let sharePointData;
    try {
      if (!SP_CHAT_URL) {
        sharePointData = {
          source: "SharePoint",
          error: "SP_CHAT_URL env var not configured."
        };
      } else {
        const spRes = await axios.post(SP_CHAT_URL, { question });
        sharePointData = { source: "SharePoint", ...spRes.data };
      }
    } catch (e) {
      sharePointData = { source: "SharePoint", error: e.message };
    }

    // -----------------------------
    // PREPARE AI INPUT
    // -----------------------------
    const combinedRaw = {
      serviceNow: serviceNowData,
      salesforce: salesforceData,
      sharePoint: sharePointData
    };

    // -----------------------------
    // AI SUMMARY (Gemini 2.0 Flash Stable)
    // -----------------------------
    let combinedAnswer = null;
    try {
      const genAI = new GoogleGenerativeAI(GEMINI_API_KEY);
      const model = genAI.getGenerativeModel({ model: "gemini-2.0-flash" });

      const prompt = `
You are generating one single executive answer based ONLY on the JSON data below.

Question:
${question}

Data:
${JSON.stringify(combinedRaw, null, 2)}

Rules:
- Make it extremely clear and leadership-friendly.
- Prioritize SALES → IT → COLLABORATION insights.
- If SharePoint has no insights, say it clearly but softly.
- If ServiceNow has incidents, quantify risk.
- If Salesforce has at-risk revenue, highlight impact.
- Avoid hallucination.
- If any system errored, mention the error politely.

Generate a final leadership summary.
`;

      const ai = await model.generateContent(prompt);
      combinedAnswer = ai.response.text();
    } catch (e) {
      combinedAnswer = null;
    }

    // -----------------------------
    // FINAL RESPONSE
    // -----------------------------
    return res.status(200).json({
      question,
      combinedAnswer: combinedAnswer || "AI model failed. See raw outputs.",
      sources: combinedRaw,
      generatedAt: new Date().toISOString()
    });

  } catch (error) {
    console.error("TXI Dashboard Fatal Error:", error);
    return res.status(500).json({
      error: "TXI Dashboard failed",
      detail: error.message
    });
  }
}
