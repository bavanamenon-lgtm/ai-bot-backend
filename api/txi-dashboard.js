// /api/txi-dashboard.js
// Vercel Serverless Function (Node)
// POST { "question": "..." }

import jsforce from "jsforce";

const JSON_HEADERS = { "Content-Type": "application/json" };

function allowCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
}

function safeNumber(n, fallback = 0) {
  const x = Number(n);
  return Number.isFinite(x) ? x : fallback;
}

function money(n) {
  const x = safeNumber(n);
  return `$${x.toLocaleString("en-US")}`;
}

function isRateLimitLike(status, text = "") {
  const t = String(text || "").toLowerCase();
  return status === 429 || status === 503 || t.includes("quota") || t.includes("resource_exhausted") || t.includes("rate");
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

async function fetchJson(url, options = {}, timeoutMs = 25000) {
  const controller = new AbortController();
  const id = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const resp = await fetch(url, { ...options, signal: controller.signal });
    const text = await resp.text();

    let json = null;
    try {
      json = JSON.parse(text);
    } catch {
      json = null;
    }

    return { ok: resp.ok, status: resp.status, text, json };
  } finally {
    clearTimeout(id);
  }
}

/** -----------------------------
 *  1) ServiceNow fetch
 *  ----------------------------- */
async function getServiceNowSummary() {
  const SN_TXI_URL = process.env.SN_TXI_URL; // https://ven06080.service-now.com/api/dtp/schindler_txi/incident_summary
  const SN_USERNAME = process.env.SN_USERNAME;
  const SN_PASSWORD = process.env.SN_PASSWORD;

  if (!SN_TXI_URL || !SN_USERNAME || !SN_PASSWORD) {
    return { source: "ServiceNow", error: "ServiceNow env vars not configured (SN_TXI_URL / SN_USERNAME / SN_PASSWORD)." };
  }

  const basic = Buffer.from(`${SN_USERNAME}:${SN_PASSWORD}`).toString("base64");
  const r = await fetchJson(
    SN_TXI_URL,
    {
      method: "GET",
      headers: {
        Authorization: `Basic ${basic}`,
        Accept: "application/json",
      },
    },
    20000
  );

  if (!r.ok) {
    return { source: "ServiceNow", error: `ServiceNow HTTP ${r.status}`, raw: r.json ?? r.text };
  }

  return r.json ?? { source: "ServiceNow", raw: r.text };
}

/** -----------------------------
 *  2) Salesforce fetch (robust)
 *  ----------------------------- */
async function getSalesforceSummary() {
  const SF_USERNAME = process.env.SF_USERNAME;
  const SF_PASSWORD = process.env.SF_PASSWORD;
  const SF_TOKEN = process.env.SF_TOKEN;
  const SF_LOGIN_URL = process.env.SF_LOGIN_URL || "https://login.salesforce.com";

  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN) {
    return { source: "Salesforce", error: "Salesforce env vars not configured (SF_USERNAME / SF_PASSWORD / SF_TOKEN)." };
  }

  const conn = new jsforce.Connection({ loginUrl: SF_LOGIN_URL });

  try {
    await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);
  } catch (e) {
    return { source: "Salesforce", error: `Salesforce login failed: ${e?.message || String(e)}` };
  }

  // Always target EBC HQ first (because you created it)
  let ebcAccount = null;

  try {
    const a = await conn.query(`SELECT Id, Name, Industry, Rating FROM Account WHERE Name = 'EBC HQ' LIMIT 1`);
    if (a?.records?.length) {
      const r = a.records[0];
      ebcAccount = { id: r.Id, name: r.Name, industry: r.Industry || null, rating: r.Rating || null };
    }
  } catch (e) {
    // ignore; fallback below
  }

  if (!ebcAccount) {
    try {
      const a2 = await conn.query(`SELECT Id, Name, Industry, Rating FROM Account WHERE Rating = 'Hot' LIMIT 1`);
      if (a2?.records?.length) {
        const r = a2.records[0];
        ebcAccount = { id: r.Id, name: r.Name, industry: r.Industry || null, rating: r.Rating || null };
      }
    } catch (e) {
      return { source: "Salesforce", error: `Account query failed: ${e?.message || String(e)}` };
    }
  }

  if (!ebcAccount) return { source: "Salesforce", error: "Could not find target account (EBC HQ / Hot account)." };

  // Try your UI custom flags first; if they don’t exist, don’t blow up the whole response.
  const oppQueries = [
    `SELECT Id, Name, Amount, StageName, CloseDate, Probability
     FROM Opportunity
     WHERE AccountId = '${ebcAccount.id}' AND At_Risk__c = true
     ORDER BY Amount DESC LIMIT 10`,
    `SELECT Id, Name, Amount, StageName, CloseDate, Probability
     FROM Opportunity
     WHERE AccountId = '${ebcAccount.id}' AND Risk_Flag__c = true
     ORDER BY Amount DESC LIMIT 10`,
    `SELECT Id, Name, Amount, StageName, CloseDate, Probability
     FROM Opportunity
     WHERE AccountId = '${ebcAccount.id}'
     ORDER BY CloseDate ASC LIMIT 10`,
  ];

  let oppRecords = [];
  let usedQueryIndex = -1;
  let lastErr = null;

  for (let i = 0; i < oppQueries.length; i++) {
    try {
      const o = await conn.query(oppQueries[i]);
      oppRecords = o?.records || [];
      usedQueryIndex = i;
      lastErr = null;
      break;
    } catch (e) {
      lastErr = e;
    }
  }

  if (usedQueryIndex === -1) {
    return { source: "Salesforce", ebcAccount, error: `Opportunity query failed: ${lastErr?.message || String(lastErr)}` };
  }

  const normalized = oppRecords.map((r) => ({
    id: r.Id,
    name: r.Name,
    amount: safeNumber(r.Amount, 0),
    stage: r.StageName,
    closeDate: r.CloseDate,
    probability: safeNumber(r.Probability, 0),
  }));

  // If we had to use fallback query, compute risk = probability <= 30 (your sample uses 10 and 30)
  const atRiskList =
    usedQueryIndex === 0 || usedQueryIndex === 1
      ? normalized
      : normalized.filter((o) => safeNumber(o.probability, 0) <= 30);

  const totalAmount = atRiskList.reduce((s, o) => s + safeNumber(o.amount), 0);

  return {
    source: "Salesforce",
    ebcAccount,
    atRiskSummary: { opportunityCount: atRiskList.length, totalAmount },
    atRiskOpportunities: atRiskList,
  };
}

/** -----------------------------
 *  3) SharePoint assistant fetch (FORCE your 3 seeded files)
 *  ----------------------------- */
async function getSharePointSummary(question) {
  const SP_CHAT_URL = process.env.SP_CHAT_URL;
  if (!SP_CHAT_URL) return { source: "SharePoint", error: "SP_CHAT_URL env var not configured." };

  // FORCE targeted retrieval every time (because leadership queries are too broad)
  const forced =
    `Leadership query. You MUST first search these exact files (by name) and summarise them if found:\n` +
    `- Annual EBC Review Notes.txt\n` +
    `- EBC_Account_Health_Risk.docx\n` +
    `- IT_Operations_Weekly_Report.docx\n\n` +
    `If you find them, return in this format:\n` +
    `1) Risk\n- Evidence (quote or short proof)\n- Customer impact\n- Action for leadership\n\n` +
    `If not found, explain EXACTLY which of the three files were not located.\n\n` +
    `User question: ${question}`;

  const r = await fetchJson(
    SP_CHAT_URL,
    { method: "POST", headers: JSON_HEADERS, body: JSON.stringify({ question: forced }) },
    30000
  );

  if (!r.ok) {
    return { source: "SharePoint", error: `SharePoint assistant HTTP ${r.status}`, raw: r.json ?? r.text };
  }

  // normalize
  return { source: "SharePoint", ...(r.json || { answer: r.text }) };
}

/** -----------------------------
 *  4) Gemini (2.5) — NEVER block output on Gemini errors
 *  ----------------------------- */
async function callGemini({ question, sources }) {
  const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
  if (!GEMINI_API_KEY) return { used: false, error: "GEMINI_API_KEY not configured." };

  // Lock to the model you want
  const model = (process.env.GEMINI_MODEL || "gemini-2.5-flash").trim();

  const instruction =
    `Write an EXECUTIVE answer. Keep it scannable.\n` +
    `Output EXACTLY:\n` +
    `EXECUTIVE BRIEF (1 line)\n` +
    `- What’s happening (max 2 bullets)\n` +
    `- Why it matters (max 2 bullets)\n` +
    `TOP 3 RISKS (numbered)\n` +
    `For each: Risk — Evidence — Action (one line each)\n` +
    `Close with: Traceability: Salesforce | ServiceNow | SharePoint\n` +
    `No long paragraphs. No filler.\n`;

  const prompt =
    `${instruction}\n\nQuestion: ${question}\n\nGround truth JSON (do not hallucinate):\n` +
    JSON.stringify(sources, null, 2);

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(
    GEMINI_API_KEY
  )}`;

  // 2 quick attempts only
  for (let attempt = 1; attempt <= 2; attempt++) {
    const r = await fetchJson(
      url,
      {
        method: "POST",
        headers: JSON_HEADERS,
        body: JSON.stringify({
          contents: [{ role: "user", parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.2, maxOutputTokens: 500 },
        }),
      },
      30000
    );

    if (r.ok && r.json) {
      const text =
        r.json?.candidates?.[0]?.content?.parts?.map((p) => p.text).join("") ||
        r.json?.candidates?.[0]?.content?.parts?.[0]?.text ||
        "";
      if (text.trim()) return { used: true, model, text };
    }

    if (isRateLimitLike(r.status, r.text)) {
      await sleep(900 * attempt);
      continue;
    }
    break;
  }

  return { used: false, error: "Gemini failed (rate/availability). Falling back to local formatting." };
}

/** -----------------------------
 *  5) Local fallback (still executive)
 *  ----------------------------- */
function buildExecutiveAnswer({ sources }) {
  const sf = sources.salesforce || {};
  const sn = sources.serviceNow || {};
  const sp = sources.sharePoint || {};

  const p1 = sn?.byPriority?.find((x) => String(x.priority) === "1")?.count ?? 0;
  const p2 = sn?.byPriority?.find((x) => String(x.priority) === "2")?.count ?? 0;

  const sfLine = sf?.atRiskSummary
    ? `Revenue risk: ${money(sf.atRiskSummary.totalAmount)} across ${sf.atRiskSummary.opportunityCount} at-risk deal(s) (${sf?.ebcAccount?.name || "key account"}).`
    : sf?.error
    ? `Revenue visibility gap: ${String(sf.error).slice(0, 140)}`
    : `Revenue: insufficient signal.`;

  const snLine = sn?.error
    ? `IT visibility gap: ${String(sn.error).slice(0, 140)}`
    : `IT stability: ${safeNumber(sn.totalHighPriority)} high-priority incidents (P1 ${safeNumber(p1)}, P2 ${safeNumber(p2)}).`;

  const spLine = sp?.error
    ? `Knowledge visibility gap: ${String(sp.error).slice(0, 140)}`
    : sp?.answer
    ? `Knowledge signal: ${String(sp.answer).slice(0, 220)}`
    : `Knowledge: no signal.`;

  return [
    `EXECUTIVE BRIEF`,
    `- What’s happening: ${snLine}`,
    `- What’s happening: ${sfLine}`,
    `- Why it matters: Leadership decisions are only as good as these signals; any blind spot is a risk itself.`,
    ``,
    `TOP 3 RISKS`,
    `1) Operations — Evidence: P1=${safeNumber(p1)} / P2=${safeNumber(p2)} — Action: 24h stabilization plan with owners + ETAs.`,
    `2) Revenue — Evidence: ${sf?.atRiskSummary ? money(sf.atRiskSummary.totalAmount) : "N/A"} — Action: exec sponsor + save-plan on EBC HQ deals.`,
    `3) Execution knowledge — Evidence: ${sp?.error ? "SharePoint failure" : "Docs signal weak/partial"} — Action: validate the 3 seeded files are searchable + returned.`,
    `Traceability: Salesforce | ServiceNow | SharePoint`,
  ].join("\n");
}

/** -----------------------------
 *  Handler
 *  ----------------------------- */
export default async function handler(req, res) {
  allowCors(res);

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: 'Use POST with JSON body { "question": "..." }' });

  let body = req.body;
  if (typeof body === "string") {
    try { body = JSON.parse(body); } catch { body = {}; }
  }

  const question = body?.question;
  if (!question || typeof question !== "string") return res.status(400).json({ error: 'Missing "question" in request body.' });

  try {
    const [serviceNow, salesforce, sharePoint] = await Promise.all([
      getServiceNowSummary(),
      getSalesforceSummary(),
      getSharePointSummary(question),
    ]);

    const sources = { serviceNow, salesforce, sharePoint };

    const gemini = await callGemini({ question, sources });
    const combinedAnswer = gemini.used ? gemini.text : buildExecutiveAnswer({ sources });

    return res.status(200).json({
      question,
      combinedAnswer,
      sources,
      gemini: gemini.used ? { used: true, model: gemini.model } : { used: false, error: gemini.error },
      generatedAt: new Date().toISOString(),
    });
  } catch (e) {
    return res.status(500).json({ error: "FUNCTION_INVOCATION_FAILED", detail: e?.message || String(e) });
  }
}
