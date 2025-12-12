// /api/txi-dashboard.js
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

/** -------------------------------------------------
 *  1) ServiceNow (standardized shape)
 *  ------------------------------------------------- */
async function getServiceNowSummary() {
  const url = process.env.SN_TXI_URL;
  const user = process.env.SN_USERNAME;
  const pass = process.env.SN_PASSWORD;

  if (!url || !user || !pass) {
    return { source: "ServiceNow", ok: false, error: "Missing SN_TXI_URL / SN_USERNAME / SN_PASSWORD", data: null };
  }

  const basic = Buffer.from(`${user}:${pass}`).toString("base64");
  const r = await fetchJson(url, {
    method: "GET",
    headers: { Authorization: `Basic ${basic}`, Accept: "application/json" }
  }, 20000);

  if (!r.ok) {
    return { source: "ServiceNow", ok: false, error: `HTTP ${r.status}`, data: r.json ?? { raw: r.text } };
  }

  // Some SN APIs wrap in {result: {...}}. Normalize.
  const payload = r.json?.result ? r.json.result : (r.json ?? null);

  return { source: "ServiceNow", ok: true, error: null, data: payload };
}

/** -------------------------------------------------
 *  2) Salesforce (no custom risk fields required)
 *  ------------------------------------------------- */
async function getSalesforceSummary() {
  const u = process.env.SF_USERNAME;
  const p = process.env.SF_PASSWORD;
  const t = process.env.SF_TOKEN;
  const loginUrl = process.env.SF_LOGIN_URL || "https://login.salesforce.com";

  if (!u || !p || !t) {
    return { source: "Salesforce", ok: false, error: "Missing SF_USERNAME / SF_PASSWORD / SF_TOKEN", data: null };
  }

  const conn = new jsforce.Connection({ loginUrl });
  try {
    await conn.login(u, p + t);
  } catch (e) {
    return { source: "Salesforce", ok: false, error: `Login failed: ${e?.message || String(e)}`, data: null };
  }

  // Account selection: prefer EBC HQ else any Hot
  let acct = null;

  try {
    const a = await conn.query(`SELECT Id, Name, Industry, Rating FROM Account WHERE Name = 'EBC HQ' LIMIT 1`);
    if (a?.records?.length) {
      const r = a.records[0];
      acct = { id: r.Id, name: r.Name, industry: r.Industry || null, rating: r.Rating || null };
    }
  } catch {}

  if (!acct) {
    try {
      const a2 = await conn.query(`SELECT Id, Name, Industry, Rating FROM Account WHERE Rating = 'Hot' LIMIT 1`);
      if (a2?.records?.length) {
        const r = a2.records[0];
        acct = { id: r.Id, name: r.Name, industry: r.Industry || null, rating: r.Rating || null };
      }
    } catch (e) {
      return { source: "Salesforce", ok: false, error: `Account query failed: ${e?.message || String(e)}`, data: null };
    }
  }

  if (!acct) {
    return { source: "Salesforce", ok: false, error: "No target account found (EBC HQ or Rating=Hot).", data: null };
  }

  // Deterministic "at-risk" criteria (portable across orgs):
  // - Not closed
  // - Probability <= 30 OR CloseDate within 45 days
  const oppQ = `
    SELECT Id, Name, Amount, StageName, CloseDate, Probability
    FROM Opportunity
    WHERE AccountId = '${acct.id}'
      AND IsClosed = false
    ORDER BY CloseDate ASC, Amount DESC
    LIMIT 25
  `;

  let opp = [];
  try {
    const o = await conn.query(oppQ);
    opp = o?.records || [];
  } catch (e) {
    return { source: "Salesforce", ok: false, error: `Opportunity query failed: ${e?.message || String(e)}`, data: { ebcAccount: acct } };
  }

  const now = new Date();
  const inDays = (dStr) => {
    const d = new Date(dStr);
    if (Number.isNaN(d.getTime())) return 99999;
    return Math.ceil((d.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
  };

  const normalized = opp.map((r) => ({
    id: r.Id,
    name: r.Name,
    amount: safeNumber(r.Amount, 0),
    stage: r.StageName,
    closeDate: r.CloseDate,
    probability: safeNumber(r.Probability, 0),
    closeInDays: r.CloseDate ? inDays(r.CloseDate) : null
  }));

  const atRisk = normalized
    .filter((o) => (o.probability <= 30) || (o.closeInDays != null && o.closeInDays <= 45))
    .sort((a, b) => (b.amount - a.amount))
    .slice(0, 10);

  const totalAmount = atRisk.reduce((s, o) => s + safeNumber(o.amount), 0);

  return {
    source: "Salesforce",
    ok: true,
    error: null,
    data: {
      ebcAccount: acct,
      atRiskSummary: { opportunityCount: atRisk.length, totalAmount },
      atRiskOpportunities: atRisk
    }
  };
}

/** -------------------------------------------------
 *  3) SharePoint assistant (keep as-is but standardized)
 *  ------------------------------------------------- */
async function getSharePointSummary(question) {
  const url = process.env.SP_CHAT_URL;
  if (!url) {
    return { source: "SharePoint", ok: false, error: "Missing SP_CHAT_URL", data: null };
  }

  const r1 = await fetchJson(url, {
    method: "POST",
    headers: JSON_HEADERS,
    body: JSON.stringify({ question })
  }, 25000);

  if (!r1.ok || !r1.json) {
    return { source: "SharePoint", ok: false, error: `HTTP ${r1.status} or non-JSON`, data: r1.json ?? { raw: r1.text } };
  }

  const ans = String(r1.json.answer || "");
  const noMatch = ans.toLowerCase().includes("couldn't find") || ans.toLowerCase().includes("no matching");

  if (!noMatch) {
    return { source: "SharePoint", ok: true, error: null, data: r1.json };
  }

  const seededPrompt =
    `Search specifically for these files and extract leadership risks + actions:\n` +
    `- Annual EBC Review Notes.txt\n` +
    `- EBC_Account_Health_Risk.docx\n` +
    `- IT_Operations_Weekly_Report.docx\n\n` +
    `Return EXACTLY 3 bullets: Risk | Customer impact | Action.\n` +
    `Original question: ${question}`;

  const r2 = await fetchJson(url, {
    method: "POST",
    headers: JSON_HEADERS,
    body: JSON.stringify({ question: seededPrompt })
  }, 25000);

  if (r2.ok && r2.json) {
    return { source: "SharePoint", ok: true, error: null, data: { ...r2.json, note: "Used seeded-file prompt fallback." } };
  }

  return {
    source: "SharePoint",
    ok: true, // assistant responded, but weak
    error: null,
    data: { ...r1.json, note: "Broad query had no match; seeded fallback failed.", seededError: r2.json ?? r2.text }
  };
}

/** -------------------------------------------------
 *  4) Executive answer (uses standardized sources.*.data)
 *  ------------------------------------------------- */
function buildExecutiveAnswer({ sources }) {
  const sf = sources.salesforce;
  const sn = sources.serviceNow;
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
  const spFirstLine = (spData.answer ? String(spData.answer).split("\n")[0] : "");
  const spLine = sp?.ok
    ? (spFirstLine ? `Knowledge signal: ${spFirstLine}` : `Knowledge signal: no explicit extract returned.`)
    : `Knowledge risk: ${String(sp?.error || "SharePoint failed").split("\n")[0]}`;

  return [
    `EXECUTIVE BRIEF — Today’s top risks`,
    `1) OPERATIONS — ${opsLine} Do now: assign owners + 24h stabilization plan (root cause, ETA).`,
    `2) REVENUE — ${salesLine} Do now: exec sponsor call + save-plan for top deals.`,
    `3) EXECUTION — ${spLine} Do now: confirm seeded docs searchability (exact name + permissions).`,
    `Traceability: Salesforce | ServiceNow | SharePoint`
  ].join("\n");
}

/** -------------------------------------------------
 *  5) Gemini (optional, never blocks)
 *  ------------------------------------------------- */
async function callGemini({ question, sources }) {
  const key = process.env.GEMINI_API_KEY;
  if (!key) return { used: false, error: "GEMINI_API_KEY not configured." };

  const model = (process.env.GEMINI_MODEL || "gemini-2.5-flash").trim();

  const instruction =
    `You are a CEO-facing executive assistant.\n` +
    `Return EXACTLY 5 lines:\n` +
    `Line1: EXECUTIVE BRIEF — <8 words>\n` +
    `Line2: 1) OPERATIONS — <1 sentence> Do now: <short>\n` +
    `Line3: 2) REVENUE — <1 sentence> Do now: <short>\n` +
    `Line4: 3) EXECUTION — <1 sentence> Do now: <short>\n` +
    `Line5: Traceability: Salesforce | ServiceNow | SharePoint\n` +
    `No extra text.\n`;

  const prompt =
    `${instruction}\n` +
    `Question: ${question}\n\n` +
    `Ground truth JSON:\n${JSON.stringify(sources, null, 2)}`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(key)}`;

  async function attemptOnce() {
    const body = {
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.2, maxOutputTokens: 400 }
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

/** -------------------------------------------------
 *  Handler
 *  ------------------------------------------------- */
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
      getSharePointSummary(question)
    ]);

    const sources = { serviceNow, salesforce, sharePoint };

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
      generatedAt: new Date().toISOString()
    });
  } catch (e) {
    return res.status(500).json({ error: "FUNCTION_INVOCATION_FAILED", detail: e?.message || String(e) });
  }
}
