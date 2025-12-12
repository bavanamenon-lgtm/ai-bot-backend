// api/txi-dashboard.js
// Vercel Serverless Function (Node 18+)
// POST /api/txi-dashboard  { "question": "..." }
//
// Env vars:
// --- ServiceNow ---
// SN_TXI_URL
// SN_USERNAME
// SN_PASSWORD
//
// --- Salesforce (jsforce login) ---
// SF_LOGIN_URL (optional; default https://login.salesforce.com)
// SF_USERNAME
// SF_PASSWORD
// SF_TOKEN
//
// --- SharePoint Assistant (your existing SP chat endpoint) ---
// SP_CHAT_URL
//
// --- Gemini (optional; never blocks output) ---
// GEMINI_API_KEY
// GEMINI_MODEL (optional; default gemini-2.5-flash)

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

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function isGeminiRateLimit(errText = "") {
  const t = String(errText || "").toLowerCase();
  return t.includes("429") || t.includes("resource_exhausted") || t.includes("quota") || t.includes("rate");
}

async function fetchJson(url, options = {}, timeoutMs = 25000) {
  const controller = new AbortController();
  const id = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const resp = await fetch(url, { ...options, signal: controller.signal });
    const text = await resp.text();

    let json = null;
    try { json = JSON.parse(text); } catch { json = null; }

    return { ok: resp.ok, status: resp.status, text, json };
  } finally {
    clearTimeout(id);
  }
}

/** -----------------------------
 *  1) ServiceNow fetch (standardized)
 *  ----------------------------- */
async function getServiceNowSummary() {
  const url = process.env.SN_TXI_URL;
  const user = process.env.SN_USERNAME;
  const pass = process.env.SN_PASSWORD;

  if (!url || !user || !pass) {
    return { source: "ServiceNow", ok: false, error: "Missing SN_TXI_URL / SN_USERNAME / SN_PASSWORD", data: null, meta: null };
  }

  const basic = Buffer.from(`${user}:${pass}`).toString("base64");

  const r = await fetchJson(
    url,
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
      ok: false,
      error: `HTTP ${r.status}`,
      data: r.json ?? { raw: r.text },
      meta: { httpStatus: r.status },
    };
  }

  // Normalize possible wrapper {result:{...}}
  const payload = r.json?.result ? r.json.result : (r.json ?? null);

  return {
    source: "ServiceNow",
    ok: true,
    error: null,
    data: payload,
    meta: { httpStatus: r.status },
  };
}

/** -----------------------------
 *  2) Salesforce fetch (portable, no custom fields)
 *  ----------------------------- */
async function getSalesforceSummary() {
  const SF_USERNAME = process.env.SF_USERNAME;
  const SF_PASSWORD = process.env.SF_PASSWORD;
  const SF_TOKEN = process.env.SF_TOKEN;
  const SF_LOGIN_URL = process.env.SF_LOGIN_URL || "https://login.salesforce.com";

  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN) {
    return { source: "Salesforce", ok: false, error: "Missing SF_USERNAME / SF_PASSWORD / SF_TOKEN", data: null, meta: null };
  }

  const conn = new jsforce.Connection({ loginUrl: SF_LOGIN_URL });

  try {
    await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);
  } catch (e) {
    return { source: "Salesforce", ok: false, error: `Login failed: ${e?.message || String(e)}`, data: null, meta: null };
  }

  let ebcAccount = null;

  // Prefer EBC HQ
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
  } catch {}

  // Fallback: any Hot account
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
      return { source: "Salesforce", ok: false, error: `Account query failed: ${e?.message || String(e)}`, data: null, meta: null };
    }
  }

  if (!ebcAccount) {
    return { source: "Salesforce", ok: false, error: "No target account found (EBC HQ or Rating=Hot).", data: null, meta: null };
  }

  // Get open opps (portable) then derive “at-risk”
  const oppQ = `
    SELECT Id, Name, Amount, StageName, CloseDate, Probability, IsClosed
    FROM Opportunity
    WHERE AccountId = '${ebcAccount.id}'
      AND IsClosed = false
    ORDER BY CloseDate ASC, Amount DESC
    LIMIT 25
  `;

  let oppRecords = [];
  try {
    const o = await conn.query(oppQ);
    oppRecords = o?.records || [];
  } catch (e) {
    return {
      source: "Salesforce",
      ok: false,
      error: `Opportunity query failed: ${e?.message || String(e)}`,
      data: { ebcAccount },
      meta: null,
    };
  }

  const now = new Date();
  const daysUntil = (dateStr) => {
    const d = new Date(dateStr);
    if (Number.isNaN(d.getTime())) return 99999;
    return Math.ceil((d.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
  };

  const normalized = oppRecords.map((r) => ({
    id: r.Id,
    name: r.Name,
    amount: safeNumber(r.Amount, 0),
    stage: r.StageName,
    closeDate: r.CloseDate,
    probability: safeNumber(r.Probability, 0),
    closeInDays: r.CloseDate ? daysUntil(r.CloseDate) : null,
  }));

  // At-risk heuristic: Probability <= 30 OR CloseDate within 45 days
  const atRisk = normalized
    .filter((o) => o.probability <= 30 || (o.closeInDays != null && o.closeInDays <= 45))
    .sort((a, b) => b.amount - a.amount)
    .slice(0, 10);

  const totalAmount = atRisk.reduce((s, o) => s + safeNumber(o.amount), 0);

  return {
    source: "Salesforce",
    ok: true,
    error: null,
    data: {
      ebcAccount,
      atRiskSummary: { opportunityCount: atRisk.length, totalAmount },
      atRiskOpportunities: atRisk,
    },
    meta: { loginUrl: SF_LOGIN_URL },
  };
}

/** -----------------------------
 *  3) SharePoint assistant fetch (standardized + seeded fallback)
 *  ----------------------------- */
async function getSharePointSummary(question) {
  const SP_CHAT_URL = process.env.SP_CHAT_URL;

  if (!SP_CHAT_URL) {
    return { source: "SharePoint", ok: false, error: "Missing SP_CHAT_URL", data: null, meta: null };
  }

  const r1 = await fetchJson(
    SP_CHAT_URL,
    { method: "POST", headers: JSON_HEADERS, body: JSON.stringify({ question }) },
    25000
  );

  if (!r1.ok || !r1.json) {
    return {
      source: "SharePoint",
      ok: false,
      error: `HTTP ${r1.status} or non-JSON`,
      data: r1.json ?? { raw: r1.text },
      meta: { httpStatus: r1.status },
    };
  }

  const ans = String(r1.json.answer || "");
  const noMatch = ans.toLowerCase().includes("couldn't find") || ans.toLowerCase().includes("no matching");

  if (!noMatch) {
    return { source: "SharePoint", ok: true, error: null, data: r1.json, meta: { httpStatus: r1.status, mode: "broad" } };
  }

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
    return {
      source: "SharePoint",
      ok: true,
      error: null,
      data: { ...r2.json, note: "Used seeded-file fallback prompt." },
      meta: { httpStatus: r2.status, mode: "seeded_fallback" },
    };
  }

  return {
    source: "SharePoint",
    ok: true,
    error: null,
    data: {
      ...r1.json,
      note: "Broad query had no match; seeded fallback failed.",
      seededError: r2.json ?? r2.text,
    },
    meta: { httpStatus: r1.status, mode: "broad_with_failed_seeded_fallback" },
  };
}

/** -----------------------------
 *  4) Deterministic executive formatter (always works)
 *  ----------------------------- */
function buildExecutiveAnswer({ sources }) {
  const sn = sources.serviceNow;
  const sf = sources.salesforce;
  const sp = sources.sharePoint;

  const snData = sn?.data || {};
  const byP = Array.isArray(snData.byPriority) ? snData.byPriority : [];
  const p1 = byP.find((x) => String(x.priority) === "1")?.count;
  const p2 = byP.find((x) => String(x.priority) === "2")?.count;

  const opsLine = sn?.ok
    ? `IT stability risk: ${safeNumber(snData.totalHighPriority)} high-priority incidents open (P1 ${safeNumber(p1)}, P2 ${safeNumber(p2)}).`
    : `IT ops visibility gap: ${String(sn?.error || "ServiceNow failed").split("\n")[0]}`;

  const sfData = sf?.data || {};
  const salesLine = sf?.ok
    ? `Revenue risk: ${sfData.atRiskSummary?.opportunityCount ?? 0} at-risk deal(s) worth ~${money(sfData.atRiskSummary?.totalAmount ?? 0)} on ${sfData.ebcAccount?.name || "a key account"}.`
    : `Revenue visibility gap: ${String(sf?.error || "Salesforce failed").split("\n")[0]}`;

  const spData = sp?.data || {};
  const spFirstLine = spData?.answer ? String(spData.answer).split("\n")[0] : "";
  const spLine = sp?.ok
    ? (spFirstLine ? `Knowledge signal: ${spFirstLine}` : `Knowledge signal: no explicit extract returned.`)
    : `Knowledge risk: ${String(sp?.error || "SharePoint failed").split("\n")[0]}`;

  const trace = `Traceability: Salesforce(${sf?.ok ? "OK" : "error"}) | ServiceNow(${sn?.ok ? "OK" : "error"}) | SharePoint(${sp?.ok ? "OK" : "error"})`;

  return [
    `EXECUTIVE BRIEF — Today’s top risks`,
    `1) OPERATIONS — ${opsLine} Do now: assign owners + 24h stabilization plan (root cause, ETA).`,
    `2) REVENUE — ${salesLine} Do now: exec sponsor call + save-plan for top deals.`,
    `3) EXECUTION — ${spLine} Do now: confirm seeded docs searchability (exact name + permissions).`,
    trace,
  ].join("\n");
}

/** -----------------------------
 *  5) Gemini (optional; never blocks)
 *  ----------------------------- */
async function callGemini({ question, sources }) {
  const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
  if (!GEMINI_API_KEY) return { used: false, error: "GEMINI_API_KEY not configured." };

  const model = (process.env.GEMINI_MODEL || "gemini-2.5-flash").trim();

  const instruction =
    `You are a CEO-facing executive assistant.\n` +
    `Return EXACTLY 5 lines:\n` +
    `Line1: EXECUTIVE BRIEF — <8 words>\n` +
    `Line2: 1) OPERATIONS — <1 sentence> Do now: <short>\n` +
    `Line3: 2) REVENUE — <1 sentence> Do now: <short>\n` +
    `Line4: 3) EXECUTION — <1 sentence> Do now: <short>\n` +
    `Line5: Traceability: Salesforce(OK/error) | ServiceNow(OK/error) | SharePoint(OK/error)\n` +
    `No extra text.\n`;

  const prompt =
    `${instruction}\n` +
    `Question: ${question}\n\n` +
    `Ground truth JSON:\n${JSON.stringify(sources, null, 2)}`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(GEMINI_API_KEY)}`;

  async function attemptOnce() {
    const body = {
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.2, maxOutputTokens: 450 },
    };

    const r = await fetchJson(url, { method: "POST", headers: JSON_HEADERS, body: JSON.stringify(body) }, 25000);
    if (!r.ok) return { ok: false, status: r.status, raw: r.json ?? r.text, errText: r.text || "" };

    const text =
      r.json?.candidates?.[0]?.content?.parts?.map((p) => p.text).join("") ||
      r.json?.candidates?.[0]?.content?.parts?.[0]?.text ||
      "";

    return { ok: true, text };
  }

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
    try { body = JSON.parse(body); } catch { body = {}; }
  }

  const question = body?.question;
  if (!question || typeof question !== "string") {
    return res.status(400).json({ error: 'Missing "question" (string) in request body.' });
  }

  try {
    const [serviceNow, salesforce, sharePoint] = await Promise.all([
      getServiceNowSummary(),
      getSalesforceSummary(),
      getSharePointSummary(question),
    ]);

    const sources = { serviceNow, salesforce, sharePoint };

    // Gemini optional — never blocks
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
