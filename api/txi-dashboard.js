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

function isGeminiRateLimit(errText = "") {
  const t = String(errText || "").toLowerCase();
  return t.includes("429") || t.includes("resource_exhausted") || t.includes("quota") || t.includes("rate");
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
  const SN_TXI_URL = process.env.SN_TXI_URL;
  const SN_USERNAME = process.env.SN_USERNAME;
  const SN_PASSWORD = process.env.SN_PASSWORD;

  if (!SN_TXI_URL || !SN_USERNAME || !SN_PASSWORD) {
    return {
      source: "ServiceNow",
      error: "ServiceNow env vars not configured (SN_TXI_URL / SN_USERNAME / SN_PASSWORD).",
    };
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
    return {
      source: "ServiceNow",
      error: `ServiceNow HTTP ${r.status}`,
      raw: r.json ?? r.text,
    };
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
    return {
      source: "Salesforce",
      error: "Salesforce env vars not configured (SF_USERNAME / SF_PASSWORD / SF_TOKEN).",
    };
  }

  const conn = new jsforce.Connection({ loginUrl: SF_LOGIN_URL });

  try {
    await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);
  } catch (e) {
    return { source: "Salesforce", error: `Salesforce login failed: ${e?.message || String(e)}` };
  }

  let ebcAccount = null;

  // Prefer the demo account you created
  try {
    const a = await conn.query(
      `SELECT Id, Name, Industry, Rating
       FROM Account
       WHERE Name = 'EBC HQ'
       LIMIT 1`
    );
    if (a?.records?.length) {
      const r = a.records[0];
      ebcAccount = { id: r.Id, name: r.Name, industry: r.Industry || null, rating: r.Rating || null };
    }
  } catch {
    // ignore
  }

  // Fallback: any Hot-rated account
  if (!ebcAccount) {
    try {
      const a2 = await conn.query(
        `SELECT Id, Name, Industry, Rating
         FROM Account
         WHERE Rating = 'Hot'
         LIMIT 1`
      );
      if (a2?.records?.length) {
        const r = a2.records[0];
        ebcAccount = { id: r.Id, name: r.Name, industry: r.Industry || null, rating: r.Rating || null };
      }
    } catch (e) {
      return { source: "Salesforce", error: `Account query failed: ${e?.message || String(e)}` };
    }
  }

  if (!ebcAccount) {
    return { source: "Salesforce", error: "Could not find a target account (EBC HQ / Hot account)." };
  }

  // Try risk fields first, else fallback by probability/close date
  const oppQueries = [
    `SELECT Id, Name, Amount, StageName, CloseDate, Probability
     FROM Opportunity
     WHERE AccountId = '${ebcAccount.id}'
       AND At_Risk__c = true
     ORDER BY Amount DESC
     LIMIT 10`,
    `SELECT Id, Name, Amount, StageName, CloseDate, Probability
     FROM Opportunity
     WHERE AccountId = '${ebcAccount.id}'
       AND Risk_Flag__c = true
     ORDER BY Amount DESC
     LIMIT 10`,
    `SELECT Id, Name, Amount, StageName, CloseDate, Probability
     FROM Opportunity
     WHERE AccountId = '${ebcAccount.id}'
     ORDER BY CloseDate ASC
     LIMIT 10`,
  ];

  let oppRecords = [];
  let lastErr = null;
  let usedFallback = false;

  for (let i = 0; i < oppQueries.length; i++) {
    try {
      const o = await conn.query(oppQueries[i]);
      oppRecords = o?.records || [];
      usedFallback = i === 2;
      lastErr = null;
      break;
    } catch (e) {
      lastErr = e;
    }
  }

  if (lastErr) {
    return {
      source: "Salesforce",
      ebcAccount,
      error: `Opportunity query failed: ${lastErr?.message || String(lastErr)}`,
    };
  }

  const normalized = oppRecords.map((r) => ({
    id: r.Id,
    name: r.Name,
    amount: safeNumber(r.Amount, 0),
    stage: r.StageName,
    closeDate: r.CloseDate,
    probability: safeNumber(r.Probability, 0),
  }));

  const listForSummary = usedFallback ? normalized.filter((o) => o.probability <= 30) : normalized;
  const totalAmount = listForSummary.reduce((s, o) => s + safeNumber(o.amount), 0);

  return {
    source: "Salesforce",
    ebcAccount,
    atRiskSummary: { opportunityCount: listForSummary.length, totalAmount },
    atRiskOpportunities: listForSummary,
  };
}

/** -----------------------------
 *  3) SharePoint assistant fetch
 *  ----------------------------- */
async function getSharePointSummary(question) {
  const SP_CHAT_URL = process.env.SP_CHAT_URL;
  if (!SP_CHAT_URL) {
    return { source: "SharePoint", error: "SP_CHAT_URL env var not configured." };
  }

  const r1 = await fetchJson(
    SP_CHAT_URL,
    { method: "POST", headers: JSON_HEADERS, body: JSON.stringify({ question }) },
    25000
  );

  if (!r1.ok) {
    return { source: "SharePoint", error: `SharePoint assistant HTTP ${r1.status}`, raw: r1.json ?? r1.text };
  }
  if (!r1.json) {
    return { source: "SharePoint", error: `SharePoint returned non-JSON`, raw: r1.text };
  }

  const ans = String(r1.json.answer || "");
  const noMatch = ans.toLowerCase().includes("couldn't find") || ans.toLowerCase().includes("no matching");

  // If no match, do a targeted second attempt for your seeded demo files
  if (noMatch) {
    const seededPrompt =
      `Search specifically for these files and extract leadership risks + actions:\n` +
      `- Annual EBC Review Notes.txt\n` +
      `- EBC_Account_Health_Risk.docx\n` +
      `- IT_Operations_Weekly_Report.docx\n\n` +
      `Return EXACTLY 3 bullets: Risk | Customer impact | Action.\n` +
      `Original question: ${question}`;

    const r2 = await fetchJson(
      SP_CHAT_URL,
      { method: "POST", headers: JSON_HEADERS, body: JSON.stringify({ question: seededPrompt }) },
      25000
    );

    if (r2.ok && r2.json) {
      return { source: "SharePoint", ...r2.json };
    }

    return {
      source: "SharePoint",
      ...r1.json, // keep the first attempt response
      note: "No match on broad query; seeded lookup attempt failed.",
      seededError: r2.json ?? r2.text,
    };
  }

  return { source: "SharePoint", ...r1.json };
}

/** -----------------------------
 *  4) Local exec formatter (always works)
 *  ----------------------------- */
function buildExecutiveAnswer({ sources }) {
  const sf = sources.salesforce || {};
  const sn = sources.serviceNow || {};
  const sp = sources.sharePoint || {};

  const p1 = sn?.byPriority?.find((x) => String(x.priority) === "1")?.count;
  const p2 = sn?.byPriority?.find((x) => String(x.priority) === "2")?.count;

  const salesLine =
    sf?.atRiskSummary
      ? `Revenue risk: ${sf.atRiskSummary.opportunityCount} at-risk deal(s) worth ~${money(sf.atRiskSummary.totalAmount)} on ${sf?.ebcAccount?.name || "a key account"}.`
      : sf?.error
      ? `Revenue visibility gap: ${String(sf.error).split("\n")[0]}`
      : `Revenue: insufficient signal.`;

  const opsLine =
    sn?.error
      ? `IT ops visibility gap: ${String(sn.error).split("\n")[0]}`
      : `IT stability risk: ${safeNumber(sn.totalHighPriority)} high-priority incidents open (P1 ${safeNumber(p1)}, P2 ${safeNumber(p2)}).`;

  const spLine =
    sp?.error
      ? `Knowledge risk: ${String(sp.error).split("\n")[0]}`
      : sp?.answer
      ? `Knowledge signal: ${String(sp.answer).split("\n")[0]}`
      : `Knowledge: no response.`;

  // STRICT “exec style”: short, do-now actions
  return [
    `EXECUTIVE BRIEF — Today’s top risks`,
    `1) OPERATIONS — ${opsLine} Do now: name owners + 24h stabilization plan (root cause, ETA).`,
    `2) REVENUE — ${salesLine} Do now: exec sponsor call + save-plan for top 2 deals.`,
    `3) EXECUTION — ${spLine} Do now: confirm the 3 seeded docs are searchable (name match + permissions).`,
    `Traceability: Salesforce | ServiceNow | SharePoint`,
  ].join("\n");
}

/** -----------------------------
 *  5) Gemini (optional) — DOES NOT BLOCK
 *  ----------------------------- */
async function callGemini({ question, sources }) {
  const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
  if (!GEMINI_API_KEY) return { used: false, error: "GEMINI_API_KEY not configured." };

  const model = (process.env.GEMINI_MODEL || "gemini-2.5-flash").trim();

  // Force short output so it’s exec-readable, but DO NOT cap tokens too low
  const instruction =
    `You are a CEO-facing executive assistant.\n` +
    `Return EXACTLY 5 lines:\n` +
    `Line1: EXECUTIVE BRIEF — <8 words>\n` +
    `Line2: 1) OPERATIONS — <1 sentence> Do now: <short>\n` +
    `Line3: 2) REVENUE — <1 sentence> Do now: <short>\n` +
    `Line4: 3) EXECUTION — <1 sentence> Do now: <short>\n` +
    `Line5: Traceability: Salesforce | ServiceNow | SharePoint\n` +
    `No extra text. No paragraphs.\n`;

  const prompt = `${instruction}\nQuestion: ${question}\n\nGround truth JSON:\n${JSON.stringify(sources, null, 2)}`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(
    GEMINI_API_KEY
  )}`;

  async function attemptOnce() {
    const body = {
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.2, maxOutputTokens: 350 },
    };

    const r = await fetchJson(url, { method: "POST", headers: JSON_HEADERS, body: JSON.stringify(body) }, 25000);

    if (!r.ok) {
      return { ok: false, status: r.status, raw: r.json ?? r.text, errText: r.text || "" };
    }

    const text =
      r.json?.candidates?.[0]?.content?.parts?.map((p) => p.text).join("") ||
      r.json?.candidates?.[0]?.content?.parts?.[0]?.text ||
      "";

    return { ok: true, text };
  }

  // Small retry for 429/503 only
  let last = null;
  for (let i = 1; i <= 2; i++) {
    const out = await attemptOnce();
    if (out.ok && out.text?.trim()) return { used: true, model, text: out.text.trim() };
    last = out;
    if (out.status === 429 || out.status === 503 || isGeminiRateLimit(out.errText)) {
      await sleep(900 * i);
      continue;
    }
    break;
  }

  return { used: false, error: `Gemini failed (HTTP ${last?.status || "?"})`, raw: last?.raw };
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
    try {
      body = JSON.parse(body);
    } catch {
      body = {};
    }
  }

  const question = body?.question;
  if (!question || typeof question !== "string") {
    return res.status(400).json({ error: 'Missing "question" in request body or not a string.' });
  }

  try {
    const [serviceNow, salesforce, sharePoint] = await Promise.all([
      getServiceNowSummary(),
      getSalesforceSummary(),
      getSharePointSummary(question),
    ]);

    const sources = { serviceNow, salesforce, sharePoint };

    // Gemini is optional: NEVER block output
    const gemini = await callGemini({ question, sources });

    const combinedAnswer =
      gemini?.used && gemini?.text
        ? gemini.text
        : buildExecutiveAnswer({ sources });

    return res.status(200).json({
      question,
      combinedAnswer,
      sources,
      gemini: gemini?.used ? { used: true, model: gemini.model } : { used: false, error: gemini?.error || null },
      generatedAt: new Date().toISOString(),
    });
  } catch (e) {
    return res.status(500).json({ error: "FUNCTION_INVOCATION_FAILED", detail: e?.message || String(e) });
  }
}
