// /api/txi-dashboard.js
// Combined leadership view across ServiceNow, Salesforce and SharePoint

import fetch from "node-fetch";
import jsforce from "jsforce";
import { GoogleGenerativeAI } from "@google/generative-ai";

const {
  SN_BASE_URL,
  SN_USER,
  SN_PASS,
  SF_USERNAME,
  SF_PASSWORD,
  SF_TOKEN,
  SF_LOGIN_URL,
  GEMINI_API_KEY,
  SP_CHAT_URL,
} = process.env;

// ---------- Helpers ----------

function ensureEnv(name) {
  if (!process.env[name]) {
    console.warn(`[txi-dashboard] Missing env var: ${name}`);
  }
}

// Check required envs once at load time (won’t break, just logs)
["SN_BASE_URL", "SN_USER", "SN_PASS", "SF_USERNAME", "SF_PASSWORD", "SF_TOKEN", "GEMINI_API_KEY"].forEach(
  ensureEnv
);

// Safe wrapper so one system failing doesn’t kill the whole response
async function safeCall(label, fn) {
  try {
    const data = await fn();
    return { source: label, ...data };
  } catch (err) {
    console.error(`[txi-dashboard] ${label} error`, err);
    return { source: label, error: err.message || String(err) };
  }
}

// ---------- ServiceNow ----------

async function fetchServiceNow() {
  if (!SN_BASE_URL || !SN_USER || !SN_PASS) {
    throw new Error("ServiceNow env vars missing (SN_BASE_URL / SN_USER / SN_PASS)");
  }

  const url = `${SN_BASE_URL.replace(/\/$/, "")}/api/dtp/schindler_txi/incident_summary`;
  const auth = Buffer.from(`${SN_USER}:${SN_PASS}`).toString("base64");

  const resp = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Basic ${auth}`,
      Accept: "application/json",
    },
  });

  if (!resp.ok) {
    const text = await resp.text().catch(() => "");
    throw new Error(`ServiceNow API error ${resp.status}: ${text}`);
  }

  const data = await resp.json();
  // Make sure there is a "source" field for consistency
  return { ...data, source: "ServiceNow" };
}

// ---------- Salesforce ----------

async function fetchSalesforce() {
  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN) {
    throw new Error("Salesforce env vars missing (SF_USERNAME / SF_PASSWORD / SF_TOKEN)");
  }

  const conn = new jsforce.Connection({
    loginUrl: SF_LOGIN_URL || "https://login.salesforce.com",
  });

  // Standard username + password + token combo
  await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);

  // 1) Get a strategic account (EBC HQ you created)
  const acctRes = await conn.query(
    "SELECT Id, Name, Industry, Rating FROM Account WHERE Name = 'EBC HQ' LIMIT 1"
  );

  const acctRec = acctRes.records && acctRes.records[0];
  const ebcAccount = acctRec
    ? {
        id: acctRec.Id,
        name: acctRec.Name,
        industry: acctRec.Industry,
        rating: acctRec.Rating,
      }
    : null;

  let atRiskOpportunities = [];
  if (ebcAccount) {
    // 2) Get opportunities for that account
    const oppRes = await conn.query(
      `SELECT Id, Name, Amount, StageName, CloseDate, Probability
       FROM Opportunity
       WHERE AccountId = '${ebcAccount.id}'`
    );

    // Treat low-probability deals as “at risk” (<= 30%)
    atRiskOpportunities = (oppRes.records || [])
      .filter((o) => typeof o.Probability === "number" && o.Probability <= 30)
      .map((o) => ({
        id: o.Id,
        name: o.Name,
        amount: o.Amount,
        stage: o.StageName,
        closeDate: o.CloseDate,
        probability: o.Probability,
      }));
  }

  const totalAmount = atRiskOpportunities.reduce(
    (sum, o) => sum + (Number(o.amount) || 0),
    0
  );

  const atRiskSummary = {
    opportunityCount: atRiskOpportunities.length,
    totalAmount,
  };

  return {
    source: "Salesforce",
    ebcAccount,
    atRiskSummary,
    atRiskOpportunities,
  };
}

// ---------- SharePoint (optional) ----------

async function fetchSharePoint(question) {
  if (!SP_CHAT_URL) {
    // Don’t break the whole flow – just report the config gap
    return {
      source: "SharePoint",
      error: "SP_CHAT_URL env var not configured.",
    };
  }

  const resp = await fetch(SP_CHAT_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ question }),
  });

  if (!resp.ok) {
    const text = await resp.text().catch(() => "");
    throw new Error(`SharePoint chat error ${resp.status}: ${text}`);
  }

  // Expecting something like { answer: "...", usedFiles: [...], candidateFiles: [...] }
  const data = await resp.json();
  return { source: "SharePoint", ...data };
}

// ---------- Gemini summarisation ----------

const genAI = GEMINI_API_KEY ? new GoogleGenerativeAI(GEMINI_API_KEY) : null;

async function generateCombinedAnswer(question, sources) {
  if (!genAI) {
    return "Gemini API key not configured. Cannot generate combined leadership view.";
  }

  const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

  const prompt = `
You are an executive assistant creating a leadership view for a CEO.

The CEO's question is:
"${question}"

You have JSON data from three systems:

ServiceNow:
${JSON.stringify(sources.serviceNow, null, 2)}

Salesforce:
${JSON.stringify(sources.salesforce, null, 2)}

SharePoint:
${JSON.stringify(sources.sharePoint, null, 2)}

Task:
- Identify today's top ~3 operational issues or risks.
- Explain clearly the *business impact* of each issue.
- Explicitly reference which system each insight came from (ServiceNow, Salesforce, SharePoint).
- Write in concise, leadership-level language, max 5 short bullet points plus a 1–2 sentence intro.
- If any system returned an error, treat that as a visibility / data-quality risk, not a technical stack trace.
`;

  const result = await model.generateContent(prompt);
  const resp = result.response;
  return resp.text();
}

// ---------- API handler ----------

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res
      .status(405)
      .json({ error: 'Use POST with JSON body { "question": "..." }' });
    return;
  }

  let question;
  try {
    ({ question } = req.body || {});
  } catch {
    // In case body parsing fails for any reason
    question = undefined;
  }

  if (typeof question !== "string" || !question.trim()) {
    res
      .status(400)
      .json({ error: 'Missing "question" in request body or not a string.' });
    return;
  }

  try {
    // Call all three systems in parallel
    const [serviceNow, salesforce, sharePoint] = await Promise.all([
      safeCall("ServiceNow", () => fetchServiceNow()),
      safeCall("Salesforce", () => fetchSalesforce()),
      safeCall("SharePoint", () => fetchSharePoint(question)),
    ]);

    const sources = { serviceNow, salesforce, sharePoint };

    const combinedAnswer = await generateCombinedAnswer(question, sources);

    res.status(200).json({
      question,
      combinedAnswer,
      sources,
      generatedAt: new Date().toISOString(),
    });
  } catch (err) {
    console.error("[txi-dashboard] Fatal error", err);
    res.status(500).json({
      error: "Failed to generate leadership dashboard answer.",
      detail: err.message || String(err),
    });
  }
}
