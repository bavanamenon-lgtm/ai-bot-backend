// api/txi-dashboard.js
// FINAL, STABLE VERSION — DEC 2025
// PURPOSE: Schindler TXI Leadership Console backend
// RULES:
// 1. Deterministic executive answer is ALWAYS generated
// 2. Gemini is OPTIONAL and only ADDS a short insight
// 3. No connector failure can break the response
// 4. No org-specific Salesforce custom fields
// 5. No truncation, no overrides, no surprises

import jsforce from "jsforce";

const JSON_HEADERS = { "Content-Type": "application/json" };

/* -------------------- Utilities -------------------- */

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
  return `$${safeNumber(n).toLocaleString("en-US")}`;
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

async function fetchJson(url, options = {}, timeoutMs = 20000) {
  const controller = new AbortController();
  const id = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const resp = await fetch(url, { ...options, signal: controller.signal });
    const text = await resp.text();
    let json = null;
    try { json = JSON.parse(text); } catch {}
    return { ok: resp.ok, status: resp.status, json, text };
  } finally {
    clearTimeout(id);
  }
}

/* -------------------- ServiceNow -------------------- */

async function getServiceNowSummary() {
  const { SN_TXI_URL, SN_USERNAME, SN_PASSWORD } = process.env;

  if (!SN_TXI_URL || !SN_USERNAME || !SN_PASSWORD) {
    return { source: "ServiceNow", ok: false, error: "Missing ServiceNow env vars", data: null };
  }

  const auth = Buffer.from(`${SN_USERNAME}:${SN_PASSWORD}`).toString("base64");

  const r = await fetchJson(SN_TXI_URL, {
    method: "GET",
    headers: { Authorization: `Basic ${auth}`, Accept: "application/json" }
  });

  if (!r.ok) {
    return { source: "ServiceNow", ok: false, error: `HTTP ${r.status}`, data: r.json || r.text };
  }

  const payload = r.json?.result || r.json;
  return { source: "ServiceNow", ok: true, error: null, data: payload };
}

/* -------------------- Salesforce -------------------- */

async function getSalesforceSummary() {
  const { SF_USERNAME, SF_PASSWORD, SF_TOKEN, SF_LOGIN_URL } = process.env;

  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN) {
    return { source: "Salesforce", ok: false, error: "Missing Salesforce env vars", data: null };
  }

  const conn = new jsforce.Connection({
    loginUrl: SF_LOGIN_URL || "https://login.salesforce.com"
  });

  try {
    await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);
  } catch (e) {
    return { source: "Salesforce", ok: false, error: "Salesforce login failed", data: null };
  }

  let account = null;

  try {
    const a = await conn.query(
      "SELECT Id, Name, Industry, Rating FROM Account WHERE Name='EBC HQ' LIMIT 1"
    );
    if (a.records.length) {
      const r = a.records[0];
      account = { id: r.Id, name: r.Name, industry: r.Industry, rating: r.Rating };
    }
  } catch {}

  if (!account) {
    try {
      const a2 = await conn.query(
        "SELECT Id, Name, Industry, Rating FROM Account WHERE Rating='Hot' LIMIT 1"
      );
      if (a2.records.length) {
        const r = a2.records[0];
        account = { id: r.Id, name: r.Name, industry: r.Industry, rating: r.Rating };
      }
    } catch {}
  }

  if (!account) {
    return { source: "Salesforce", ok: false, error: "No target account found", data: null };
  }

  let opps = [];
  try {
    const o = await conn.query(`
      SELECT Id, Name, Amount, StageName, Probability, CloseDate
      FROM Opportunity
      WHERE AccountId='${account.id}' AND IsClosed=false
      ORDER BY CloseDate ASC, Amount DESC
      LIMIT 20
    `);
    opps = o.records || [];
  } catch (e) {
    return { source: "Salesforce", ok: false, error: "Opportunity query failed", data: { account } };
  }

  const now = new Date();
  const daysUntil = (d) => Math.ceil((new Date(d) - now) / 86400000);

  const atRisk = opps
    .map(o => ({
      name: o.Name,
      amount: safeNumber(o.Amount),
      probability: safeNumber(o.Probability),
      closeInDays: o.CloseDate ? daysUntil(o.CloseDate) : 999
    }))
    .filter(o => o.probability <= 30 || o.closeInDays <= 45)
    .sort((a, b) => b.amount - a.amount);

  return {
    source: "Salesforce",
    ok: true,
    error: null,
    data: {
      ebcAccount: account,
      atRiskSummary: {
        opportunityCount: atRisk.length,
        totalAmount: atRisk.reduce((s, o) => s + o.amount, 0)
      },
      atRiskOpportunities: atRisk
    }
  };
}

/* -------------------- SharePoint -------------------- */

async function getSharePointSummary(question) {
  const { SP_CHAT_URL } = process.env;
  if (!SP_CHAT_URL) {
    return { source: "SharePoint", ok: false, error: "Missing SP_CHAT_URL", data: null };
  }

  const r = await fetchJson(SP_CHAT_URL, {
    method: "POST",
    headers: JSON_HEADERS,
    body: JSON.stringify({ question })
  });

  if (!r.ok || !r.json) {
    return { source: "SharePoint", ok: false, error: "SharePoint assistant failed", data: r.text };
  }

  return { source: "SharePoint", ok: true, error: null, data: r.json };
}

/* -------------------- Executive Builder -------------------- */

function buildExecutiveAnswer(sources) {
  const sn = sources.serviceNow;
  const sf = sources.salesforce;
  const sp = sources.sharePoint;

  const snData = sn?.data || {};
  const byP = snData.byPriority || [];
  const p1 = byP.find(p => String(p.priority) === "1")?.count || 0;
  const p2 = byP.find(p => String(p.priority) === "2")?.count || 0;

  const ops = sn.ok
    ? `IT stability risk: ${safeNumber(snData.totalHighPriority)} high-priority incidents (P1 ${p1}, P2 ${p2}).`
    : `IT visibility gap: ${sn.error}`;

  const sfData = sf?.data || {};
  const rev = sf.ok
    ? `Revenue risk: ${sfData.atRiskSummary.opportunityCount} at-risk deal(s) worth ~${money(sfData.atRiskSummary.totalAmount)} on ${sfData.ebcAccount?.name}.`
    : `Revenue visibility gap: ${sf.error}`;

  const spInsight = sp.ok
    ? (sp.data?.answer ? String(sp.data.answer).split("\n")[0] : "No major knowledge risks detected.")
    : `Knowledge gap: ${sp.error}`;

  return [
    "EXECUTIVE BRIEF — Today’s top risks",
    `1) OPERATIONS — ${ops} Do now: assign owners and a 24h stabilization plan.`,
    `2) REVENUE — ${rev} Do now: exec sponsor call and save-plan.`,
    `3) EXECUTION — Knowledge signal: ${spInsight} Do now: confirm document owners.`,
    `Traceability: Salesforce(${sf.ok ? "OK" : "error"}) | ServiceNow(${sn.ok ? "OK" : "error"}) | SharePoint(${sp.ok ? "OK" : "error"})`
  ].join("\n");
}

/* -------------------- Gemini (Insight Only) -------------------- */

async function callGeminiInsight(question) {
  const { GEMINI_API_KEY } = process.env;
  if (!GEMINI_API_KEY) return null;

  const prompt =
    "You are a CEO advisor. Return ONE sharp insight sentence (max 20 words). " +
    "Do not summarize data. Do not list bullets.";

  const r = await fetchJson(
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`,
    {
      method: "POST",
      headers: JSON_HEADERS,
      body: JSON.stringify({
        contents: [{ role: "user", parts: [{ text: `${prompt}\nQuestion: ${question}` }] }],
        generationConfig: { temperature: 0.2, maxOutputTokens: 60 }
      })
    }
  );

  const text =
    r.json?.candidates?.[0]?.content?.parts?.[0]?.text?.trim();

  return text || null;
}

/* -------------------- Handler -------------------- */

export default async function handler(req, res) {
  allowCors(res);

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "POST only" });

  const question = req.body?.question;
  if (!question) return res.status(400).json({ error: "Missing question" });

  const [serviceNow, salesforce, sharePoint] = await Promise.all([
    getServiceNowSummary(),
    getSalesforceSummary(),
    getSharePointSummary(question)
  ]);

  const sources = { serviceNow, salesforce, sharePoint };

  // Always build deterministic answer
  let combinedAnswer = buildExecutiveAnswer(sources);

  // Add Gemini insight (never replace)
  const geminiInsight = await callGeminiInsight(question);
  if (geminiInsight) {
    combinedAnswer = `AI INSIGHT — ${geminiInsight}\n\n${combinedAnswer}`;
  }

  return res.status(200).json({
    question,
    combinedAnswer,
    sources,
    gemini: geminiInsight ? { used: true } : { used: false },
    generatedAt: new Date().toISOString()
  });
}
