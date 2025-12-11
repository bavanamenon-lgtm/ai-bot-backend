// api/txi-dashboard.js
// Combined leadership view across ServiceNow, Salesforce, SharePoint

import fetch from "node-fetch";
import jsforce from "jsforce";

const {
  SN_BASE_URL,
  SN_USER,
  SN_PASS,
  SF_USERNAME,
  SF_PASSWORD,
  SF_TOKEN,
  SF_LOGIN_URL,
  GEMINI_API_KEY,
} = process.env;

// The single leadership question we are targeting
const LEADERSHIP_QUESTION =
  "What are the top 3 operational issues I should care about today, and what’s the business impact?";

// ---------- Helper: call ServiceNow Scripted REST ----------

async function getServiceNowSummary() {
  if (!SN_BASE_URL || !SN_USER || !SN_PASS) {
    throw new Error("ServiceNow credentials not set");
  }

  const url = `${SN_BASE_URL}/api/schindler_txi/incident_summary`;

  const auth = Buffer.from(`${SN_USER}:${SN_PASS}`).toString("base64");

  const resp = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Basic ${auth}`,
      Accept: "application/json",
    },
  });

  const text = await resp.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    throw new Error("ServiceNow returned non-JSON");
  }

  if (!resp.ok) {
    throw new Error("ServiceNow API error");
  }

  return data; // { source, totalHighPriority, byPriority, ebcIncidents[] }
}

// ---------- Helper: call Salesforce via jsforce ----------

async function getSalesforceSummary() {
  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN) {
    throw new Error("Salesforce credentials not set");
  }

  const conn = new jsforce.Connection({
    loginUrl: SF_LOGIN_URL || "https://login.salesforce.com",
  });

  await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);

  // 1) Pull EBC account
  const accRes = await conn.query(
    "SELECT Id, Name, Industry, Rating, Type, AnnualRevenue " +
      "FROM Account " +
      "WHERE Name LIKE '%EBC Elevators AG%' " +
      "LIMIT 1"
  );

  const ebcAccount = accRes.records[0] || null;
  let opps = [];
  if (ebcAccount) {
    const oppRes = await conn.query(
      "SELECT Id, Name, StageName, Amount, CloseDate, Probability " +
        "FROM Opportunity " +
        "WHERE AccountId = '" +
        ebcAccount.Id +
        "'" +
        " ORDER BY CloseDate DESC " +
        " LIMIT 10"
    );
    opps = oppRes.records || [];
  }

  // Simple derived view for POC
  const atRiskOpportunities = opps.filter((o) => {
    const prob = Number(o.Probability || 0);
    // "At risk" heuristic: high value but probability < 60
    const amt = Number(o.Amount || 0);
    return amt >= 50000 && prob < 60;
  });

  return {
    source: "Salesforce",
    ebcAccount: ebcAccount
      ? {
          id: ebcAccount.Id,
          name: ebcAccount.Name,
          industry: ebcAccount.Industry,
          rating: ebcAccount.Rating,
          type: ebcAccount.Type,
          annualRevenue: ebcAccount.AnnualRevenue,
        }
      : null,
    atRiskOpportunities: atRiskOpportunities.map((o) => ({
      id: o.Id,
      name: o.Name,
      stageName: o.StageName,
      amount: o.Amount,
      closeDate: o.CloseDate,
      probability: o.Probability,
    })),
  };
}

// ---------- Helper: call SharePoint assistant we already built ----------

async function getSharePointSummary() {
  // We will re-use /api/chat-sp with a fixed question
  const question =
    "Search recent documents related to 'EBC' and summarise key operational risks, customer escalations, deadlines, and commitments.";

  // Since this function is in the same Vercel project, we call via relative URL in production.
  // For simplicity, we just call the same host path. Vercel will route it correctly.
  const url = `${process.env.VERCEL_URL
    ? "https://" + process.env.VERCEL_URL
    : ""}/api/chat-sp`;

  const payload = { question };

  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  const text = await resp.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    throw new Error("SharePoint /api/chat-sp returned non-JSON");
  }

  if (!resp.ok) {
    throw new Error("SharePoint assistant API error");
  }

  // We only need the answer text for leadership
  return {
    source: "SharePoint",
    raw: data,
    summary: data.answer || "",
  };
}

// ---------- Helper: call Gemini for combined leadership answer ----------

async function callGeminiCombined({ sn, sf, sp }) {
  if (!GEMINI_API_KEY) {
    throw new Error("GEMINI_API_KEY is not set");
  }

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" +
    GEMINI_API_KEY;

  const prompt = `
You are helping a C-level leader at Schindler.

Question:
${LEADERSHIP_QUESTION}

You are given three data sources in JSON form:

1) ServiceNow incidents:
${JSON.stringify(sn, null, 2)}

2) Salesforce accounts & opportunities:
${JSON.stringify(sf, null, 2)}

3) SharePoint documents summary:
${JSON.stringify(sp, null, 2)}

Rules:
- Give a concise, leadership-level answer.
- Start with a short paragraph: 2–3 sentences summarising the situation.
- Then list the **top 3 operational issues** as bullet points.
- For each issue, mention:
  - Which system(s) it comes from (ServiceNow / Salesforce / SharePoint)
  - The business impact in plain language
- Do NOT expose IDs or internal field names.
- If something is unclear, say it is based on limited sample data (this is a POC).
`.trim();

  const body = {
    contents: [{ parts: [{ text: prompt }] }],
  };

  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  const raw = await resp.text();
  let data;
  try {
    data = JSON.parse(raw);
  } catch (e) {
    throw new Error("Gemini non-JSON response");
  }

  if (!resp.ok) {
    throw new Error("Gemini API error");
  }

  const candidate = data.candidates && data.candidates[0];
  const parts = candidate && candidate.content && candidate.content.parts;
  const answer = parts && parts[0] && parts[0].text;

  return answer || "No combined answer generated.";
}

// ---------- Main handler ----------

export default async function handler(req, res) {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Use GET" });
    return;
  }

  try {
    const [sn, sf, sp] = await Promise.all([
      getServiceNowSummary().catch((e) => ({
        error: String(e),
        source: "ServiceNow",
      })),
      getSalesforceSummary().catch((e) => ({
        error: String(e),
        source: "Salesforce",
      })),
      getSharePointSummary().catch((e) => ({
        error: String(e),
        source: "SharePoint",
      })),
    ]);

    let combinedAnswer = "";
    try {
      combinedAnswer = await callGeminiCombined({ sn, sf, sp });
    } catch (e) {
      combinedAnswer =
        "I could not generate a combined AI answer due to an internal error, but here are the raw summaries.";
    }

    res.status(200).json({
      question: LEADERSHIP_QUESTION,
      combinedAnswer,
      sources: {
        serviceNow: sn,
        salesforce: sf,
        sharePoint: sp,
      },
      generatedAt: new Date().toISOString(),
    });
  } catch (err) {
    console.error("Error in /api/txi-dashboard:", String(err));
    res.status(500).json({
      error: "Internal error in /api/txi-dashboard",
      details: String(err),
    });
  }
}
