// api/txi-dashboard.js
// TXI POC Master Endpoint
// POST /api/txi-dashboard { "question": "..." }
//
// Key: Executive Response Contract enforced.
// SharePoint: direct Graph read via /api/sharepoint-signals (same Vercel deployment)
//
// Env:
// SN_TXI_URL, SN_USERNAME, SN_PASSWORD
// SF_USERNAME, SF_PASSWORD, SF_TOKEN, SF_LOGIN_URL(optional)
// MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET (for SharePoint Graph)
// Optional: GEMINI_API_KEY, GEMINI_MODEL

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
  const x = safeNumber(n, 0);
  return `$${x.toLocaleString("en-US")}`;
}

async function fetchJson(url, options = {}, timeoutMs = 25000) {
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

/* ----------------------------- ServiceNow ----------------------------- */
async function getServiceNowSummary() {
  const url = process.env.SN_TXI_URL;
  const user = process.env.SN_USERNAME;
  const pass = process.env.SN_PASSWORD;

  if (!url || !user || !pass) {
    return { source: "ServiceNow", ok: false, error: "Missing SN env vars", data: null };
  }

  const basic = Buffer.from(`${user}:${pass}`).toString("base64");
  const r = await fetchJson(url, {
    method: "GET",
    headers: { Authorization: `Basic ${basic}`, Accept: "application/json" }
  }, 20000);

  if (!r.ok) return { source: "ServiceNow", ok: false, error: `HTTP ${r.status}`, data: r.json ?? { raw: r.text } };

  const payload = r.json?.result ? r.json.result : (r.json ?? null);
  return { source: "ServiceNow", ok: true, error: null, data: payload };
}

/* ----------------------------- Salesforce ----------------------------- */
async function getSalesforceSummary() {
  const SF_USERNAME = process.env.SF_USERNAME;
  const SF_PASSWORD = process.env.SF_PASSWORD;
  const SF_TOKEN = process.env.SF_TOKEN;
  const SF_LOGIN_URL = process.env.SF_LOGIN_URL || "https://login.salesforce.com";

  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN) {
    return { source: "Salesforce", ok: false, error: "Missing SF env vars", data: null };
  }

  const conn = new jsforce.Connection({ loginUrl: SF_LOGIN_URL });

  try {
    await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);
  } catch (e) {
    return { source: "Salesforce", ok: false, error: `Login failed: ${e?.message || String(e)}`, data: null };
  }

  let acct = null;
  try {
    const a = await conn.query(`SELECT Id, Name, Industry, Rating FROM Account WHERE Name='EBC HQ' LIMIT 1`);
    if (a?.records?.length) {
      const r = a.records[0];
      acct = { id: r.Id, name: r.Name, industry: r.Industry || null, rating: r.Rating || null };
    }
  } catch {}

  if (!acct) {
    try {
      const a2 = await conn.query(`SELECT Id, Name, Industry, Rating FROM Account WHERE Rating='Hot' LIMIT 1`);
      if (a2?.records?.length) {
        const r = a2.records[0];
        acct = { id: r.Id, name: r.Name, industry: r.Industry || null, rating: r.Rating || null };
      }
    } catch (e) {
      return { source: "Salesforce", ok: false, error: `Account query failed: ${e?.message || String(e)}`, data: null };
    }
  }

  if (!acct) return { source: "Salesforce", ok: false, error: "No target account found", data: null };

  let oppRecords = [];
  try {
    const o = await conn.query(`
      SELECT Id, Name, Amount, StageName, CloseDate, Probability, IsClosed
      FROM Opportunity
      WHERE AccountId='${acct.id}' AND IsClosed=false
      ORDER BY CloseDate ASC, Amount DESC
      LIMIT 25
    `);
    oppRecords = o?.records || [];
  } catch (e) {
    return { source: "Salesforce", ok: false, error: `Opportunity query failed: ${e?.message || String(e)}`, data: { ebcAccount: acct } };
  }

  const now = new Date();
  const daysUntil = (dStr) => {
    const d = new Date(dStr);
    if (Number.isNaN(d.getTime())) return 99999;
    return Math.ceil((d.getTime() - now.getTime()) / 86400000);
  };

  const normalized = oppRecords.map((r) => ({
    id: r.Id,
    name: r.Name,
    amount: safeNumber(r.Amount, 0),
    stage: r.StageName,
    closeDate: r.CloseDate,
    probability: safeNumber(r.Probability, 0),
    closeInDays: r.CloseDate ? daysUntil(r.CloseDate) : null
  }));

  const atRisk = normalized
    .filter((o) => (o.probability <= 30) || (o.closeInDays != null && o.closeInDays <= 45))
    .sort((a, b) => b.amount - a.amount)
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

/* ----------------------------- SharePoint (Graph via internal endpoint) ----------------------------- */
async function getSharePointSignals(question, req) {
  // Call same deployment endpoint
  // Works on Vercel because it routes internally over HTTPS using host header
  const proto = (req.headers["x-forwarded-proto"] || "https");
  const host = req.headers["x-forwarded-host"] || req.headers.host;
  const base = `${proto}://${host}`;

  const r = await fetchJson(`${base}/api/sharepoint-signals`, {
    method: "POST",
    headers: JSON_HEADERS,
    body: JSON.stringify({ question })
  }, 30000);

  // sharepoint-signals returns 200 always with ok true/false in payload
  const payload = r.json || { ok: false, error: "SharePoint signals returned non-JSON" };

  return {
    source: "SharePoint",
    ok: !!payload.ok,
    error: payload.ok ? null : (payload.error || "NO_MATCH"),
    data: payload
  };
}

/* ----------------------------- Executive Response Contract ----------------------------- */

function computeRisk(sn, sf, sp) {
  // Forced clarity: High/Medium/Low only
  const snData = sn?.data || {};
  const byP = Array.isArray(snData.byPriority) ? snData.byPriority : [];
  const p1 = safeNumber(byP.find((x) => String(x.priority) === "1")?.count, 0);
  const totalHP = safeNumber(snData.totalHighPriority, 0);

  const sfData = sf?.data || {};
  const rev = safeNumber(sfData.atRiskSummary?.totalAmount, 0);

  if (!sn?.ok) return "High";
  if (p1 >= 50 || totalHP >= 75) return "High";
  if (rev >= 250000) return "Medium";
  if (!sp?.ok) return "Medium";
  return "Low";
}

function noHedge(text) {
  // kill hedge words if they appear accidentally
  return String(text || "")
    .replace(/\b(might|could|possibly|maybe|likely|potentially)\b/gi, "")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function buildExecBriefContract({ question, sources }) {
  const sn = sources.serviceNow;
  const sf = sources.salesforce;
  const sp = sources.sharePoint;

  const snData = sn?.data || {};
  const byP = Array.isArray(snData.byPriority) ? snData.byPriority : [];
  const p1 = safeNumber(byP.find((x) => String(x.priority) === "1")?.count, 0);
  const p2 = safeNumber(byP.find((x) => String(x.priority) === "2")?.count, 0);
  const totalHP = safeNumber(snData.totalHighPriority, 0);

  const sfData = sf?.data || {};
  const acct = sfData.ebcAccount?.name || "a key account";
  const industry = sfData.ebcAccount?.industry || "—";
  const dealCount = safeNumber(sfData.atRiskSummary?.opportunityCount, 0);
  const dealValue = safeNumber(sfData.atRiskSummary?.totalAmount, 0);

  const spData = sp?.data || {};
  const hasDocs = sp?.ok && (spData?.filesFound?.length > 0);
  const knowledgeLine = hasDocs
    ? "Leadership notes are available to confirm impacted areas and priority customers."
    : "Leadership notes are not visible, which blocks precise impact confirmation.";

  const riskLevel = computeRisk(sn, sf, sp);

  const s1 = sn?.ok
    ? `High-severity service disruption is above baseline today (${totalHP} high-priority issues; P1 ${p1}, P2 ${p2}).`
    : `Operational disruption is elevated, but live visibility is degraded right now.`;

  const s2 = (sf?.ok && dealCount > 0)
    ? `This threatens customer experience today and puts ${dealCount} active deal(s) worth ~${money(dealValue)} at risk if not contained.`
    : `This threatens customer experience today and requires immediate containment to protect service commitments.`;

  const s3 = (sf?.ok)
    ? `Primary commercial exposure sits with ${acct} (${industry}); operational impact concentrates where the highest-severity issues are open.`
    : `Impact concentrates where the highest-severity issues are open and where commercial commitments are time-sensitive.`;

  const s4 = `Risk: ${riskLevel}.`;

  const s5 = (riskLevel === "High")
    ? `Leadership action: assign a single incident commander now, lock a 24-hour stabilization plan, and trigger proactive customer communication for priority accounts.`
    : `Leadership action: confirm owners and timelines today, protect priority deals with executive outreach, and restore leadership note visibility for faster decisions.`;

  const out = [
    "EXECUTIVE BRIEF — Today’s Primary Risk & Customer Impact",
    "",
    "1. What’s happening",
    noHedge(s1),
    "",
    "2. Why this matters",
    noHedge(s2),
    "",
    "3. Who is impacted",
    noHedge(`${s3} ${knowledgeLine}`),
    "",
    "4. Risk level",
    noHedge(s4),
    "",
    "5. Leadership attention required",
    noHedge(s5)
  ].join("\n");

  return out;
}

// Reject Gemini if it violates contract
function violatesContract(text) {
  const t = String(text || "");

  // Must have max 5 sections + these headings (exact)
  const must = [
    "EXECUTIVE BRIEF — Today’s Primary Risk & Customer Impact",
    "1. What’s happening",
    "2. Why this matters",
    "3. Who is impacted",
    "4. Risk level",
    "5. Leadership attention required"
  ];
  if (!must.every(m => t.includes(m))) return true;

  // No system names / technical words
  const forbidden = /(servicenow|salesforce|sharepoint|api|http|token|drive|site id|soql|endpoint|graph)/i;
  if (forbidden.test(t)) return true;

  // No hedge words
  const hedge = /\b(might|could|possibly|maybe|likely|potentially)\b/i;
  if (hedge.test(t)) return true;

  // Risk level must be explicit High/Medium/Low
  if (!/Risk:\s*(High|Medium|Low)\b/.test(t)) return true;

  // Each section should be 1–2 sentences. (Approx check: limit per section lines)
  // We enforce by limiting total length and expecting compact structure.
  if (t.length > 1100) return true;
  if (t.length < 350) return true;

  return false;
}

async function callGeminiExec(question, contextSignals) {
  const key = process.env.GEMINI_API_KEY;
  if (!key) return { used: false, error: "GEMINI_API_KEY not configured." };

  const model = (process.env.GEMINI_MODEL || "gemini-2.5-flash").trim();

  // EXECUTIVE RESPONSE CONTRACT embedded (your contract)
  const systemContract =
`You are an enterprise executive AI assistant acting as Chief of Staff to C-level leadership.
Optimize for decision clarity, not completeness.

Strict rules:
- Use plain business language (no system names, no technical details).
- Maximum 5 sections.
- Each section: 1–2 short sentences.
- No hedging words (might, could, possibly).
- Explicitly state risk level (High / Medium / Low).
- End with leadership actions, not analysis.
- Do NOT explain reasoning.
- Do NOT mention data sources unless asked.

You must follow this structure exactly:

EXECUTIVE BRIEF — Today’s Primary Risk & Customer Impact
1. What’s happening
2. Why this matters
3. Who is impacted
4. Risk level
5. Leadership attention required`;

  const userPrompt =
`Create an EXECUTIVE BRIEF for C-level leadership answering the question below.

Question:
"${question}"

Context signals (synthesize, do not repeat verbatim):
${contextSignals}

Constraints:
- Follow the Executive Brief structure
- Maximum 5 sections
- Each section 1–2 short sentences
- Use plain business language only
- Explicitly state risk level
- End with leadership actions
- Do not include system names or technical steps`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(key)}`;

  const body = {
    contents: [{ role: "user", parts: [{ text: `${systemContract}\n\n${userPrompt}` }] }],
    generationConfig: { temperature: 0.2, maxOutputTokens: 650 }
  };

  const r = await fetchJson(url, { method: "POST", headers: JSON_HEADERS, body: JSON.stringify(body) }, 25000);
  if (!r.ok) return { used: false, error: `Gemini HTTP ${r.status}`, raw: r.json ?? r.text };

  const text =
    r.json?.candidates?.[0]?.content?.parts?.map((p) => p.text).join("") ||
    r.json?.candidates?.[0]?.content?.parts?.[0]?.text ||
    "";

  return { used: true, model, text: String(text || "").trim() };
}

/* ----------------------------- Handler ----------------------------- */
export default async function handler(req, res) {
  allowCors(res);

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "POST only" });

  let body = req.body;
  if (typeof body === "string") {
    try { body = JSON.parse(body); } catch { body = {}; }
  }

  const question = body?.question;
  if (!question || typeof question !== "string") {
    return res.status(400).json({ error: 'Missing "question" (string).' });
  }

  try {
    const [serviceNow, salesforce, sharePoint] = await Promise.all([
      getServiceNowSummary(),
      getSalesforceSummary(),
      getSharePointSignals(question, req)
    ]);

    const sources = { serviceNow, salesforce, sharePoint };

    // Deterministic contract answer always exists
    const deterministic = buildExecBriefContract({ question, sources });

    // Build context signals for Gemini (still plain language, no system names)
    const snData = serviceNow?.data || {};
    const byP = Array.isArray(snData.byPriority) ? snData.byPriority : [];
    const p1 = safeNumber(byP.find((x) => String(x.priority) === "1")?.count, 0);
    const p2 = safeNumber(byP.find((x) => String(x.priority) === "2")?.count, 0);
    const totalHP = safeNumber(snData.totalHighPriority, 0);

    const sfData = salesforce?.data || {};
    const acct = sfData.ebcAccount?.name || "a key account";
    const industry = sfData.ebcAccount?.industry || "—";
    const dealCount = safeNumber(sfData.atRiskSummary?.opportunityCount, 0);
    const dealValue = safeNumber(sfData.atRiskSummary?.totalAmount, 0);

    const knowledgeGap = sharePoint?.ok ? "Leadership notes available." : "Knowledge visibility gaps affecting impact assessment.";

    const contextSignals =
`- High-priority issues: P1=${p1}, P2=${p2}, total=${totalHP} (above normal baseline)
- Revenue exposure: ${dealCount} active deal(s), ~${money(dealValue)}
- Key account: ${acct} (${industry})
- ${knowledgeGap}`;

    // Gemini optional: accept only if it respects contract
    let combinedAnswer = deterministic;
    let geminiMeta = { used: false, error: "Gemini not used." };

    const g = await callGeminiExec(question, contextSignals);
    if (g.used && g.text && !violatesContract(g.text)) {
      combinedAnswer = g.text;
      geminiMeta = { used: true, model: g.model };
    } else if (g.used) {
      geminiMeta = { used: false, error: "Gemini output rejected (contract violation)." };
    }

    return res.status(200).json({
      question,
      combinedAnswer,
      sources,
      gemini: geminiMeta,
      generatedAt: new Date().toISOString()
    });
  } catch (e) {
    return res.status(500).json({ error: "FUNCTION_INVOCATION_FAILED", detail: e?.message || String(e) });
  }
}
