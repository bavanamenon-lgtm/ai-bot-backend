// api/txi-dashboard.js
// Schindler Total Intelligence – Leadership Console (TXI POC)
// FINAL: Canonical Exec Template + SharePoint "no match" treated as NOT OK
//
// POST /api/txi-dashboard { "question": "..." }
//
// Env vars:
// SN_TXI_URL, SN_USERNAME, SN_PASSWORD
// SF_USERNAME, SF_PASSWORD, SF_TOKEN, SF_LOGIN_URL (optional)
// SP_CHAT_URL
// GEMINI_API_KEY (optional), GEMINI_MODEL (optional; default gemini-2.5-flash)

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

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function isNoMatchText(s = "") {
  const t = String(s || "").toLowerCase();
  return (
    t.includes("couldn't find") ||
    t.includes("could not find") ||
    t.includes("no matching") ||
    t.includes("i couldn't find any matching") ||
    t.includes("try using an exact file name")
  );
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

/* ----------------------------- 1) ServiceNow ----------------------------- */
async function getServiceNowSummary() {
  const url = process.env.SN_TXI_URL;
  const user = process.env.SN_USERNAME;
  const pass = process.env.SN_PASSWORD;

  if (!url || !user || !pass) {
    return { source: "ServiceNow", ok: false, error: "Missing SN_TXI_URL / SN_USERNAME / SN_PASSWORD", data: null };
  }

  const basic = Buffer.from(`${user}:${pass}`).toString("base64");

  const r = await fetchJson(
    url,
    { method: "GET", headers: { Authorization: `Basic ${basic}`, Accept: "application/json" } },
    20000
  );

  if (!r.ok) {
    return { source: "ServiceNow", ok: false, error: `HTTP ${r.status}`, data: r.json ?? { raw: r.text } };
  }

  const payload = r.json?.result ? r.json.result : (r.json ?? null);
  return { source: "ServiceNow", ok: true, error: null, data: payload };
}

/* ----------------------------- 2) Salesforce ----------------------------- */
async function getSalesforceSummary() {
  const SF_USERNAME = process.env.SF_USERNAME;
  const SF_PASSWORD = process.env.SF_PASSWORD;
  const SF_TOKEN = process.env.SF_TOKEN;
  const SF_LOGIN_URL = process.env.SF_LOGIN_URL || "https://login.salesforce.com";

  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN) {
    return { source: "Salesforce", ok: false, error: "Missing SF_USERNAME / SF_PASSWORD / SF_TOKEN", data: null };
  }

  const conn = new jsforce.Connection({ loginUrl: SF_LOGIN_URL });

  try {
    await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);
  } catch (e) {
    return { source: "Salesforce", ok: false, error: `Login failed: ${e?.message || String(e)}`, data: null };
  }

  // Prefer EBC HQ, else any Hot account
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

  if (!acct) {
    return { source: "Salesforce", ok: false, error: "No target account found (EBC HQ or Rating=Hot).", data: null };
  }

  // Portable at-risk heuristic (no custom fields):
  // - Not closed
  // - Probability <= 30 OR CloseDate <= 45 days
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
    return Math.ceil((d.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
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

/* ----------------------------- 3) SharePoint ----------------------------- */
async function askSharePointAssistant(prompt) {
  const SP_CHAT_URL = process.env.SP_CHAT_URL;
  if (!SP_CHAT_URL) {
    return { ok: false, status: 0, json: null, text: "Missing SP_CHAT_URL" };
  }
  return fetchJson(
    SP_CHAT_URL,
    { method: "POST", headers: JSON_HEADERS, body: JSON.stringify({ question: prompt }) },
    25000
  );
}

async function getSharePointSummary(question) {
  const SP_CHAT_URL = process.env.SP_CHAT_URL;
  if (!SP_CHAT_URL) {
    return { source: "SharePoint", ok: false, error: "Missing SP_CHAT_URL", data: null };
  }

  // Seeded-first (because your library needs exact filenames to work reliably)
  const seededPrompt =
    `Search ONLY these seeded files and extract leadership signals:\n` +
    `1) Annual EBC Review Notes.txt\n` +
    `2) EBC_Account_Health_Risk.docx\n` +
    `3) IT_Operations_Weekly_Report.docx\n\n` +
    `Return:\n` +
    `- One line called: "Key risk signals:"\n` +
    `- Then 3 bullets max.\n` +
    `- Each bullet format: Risk | Customer impact | Suggested action\n\n` +
    `If you still cannot find them, say exactly: "NO_MATCH".\n\n` +
    `User question context: ${question}`;

  const rSeed = await askSharePointAssistant(seededPrompt);
  if (!rSeed.ok || !rSeed.json) {
    return { source: "SharePoint", ok: false, error: `HTTP ${rSeed.status} or non-JSON`, data: rSeed.json ?? { raw: rSeed.text } };
  }

  const seededAnswer = String(rSeed.json.answer || "");
  if (!seededAnswer || seededAnswer.trim() === "NO_MATCH" || isNoMatchText(seededAnswer)) {
    // Fallback to the user's question (broad), but treat "no match" as NOT OK
    const rBroad = await askSharePointAssistant(question);

    if (!rBroad.ok || !rBroad.json) {
      return { source: "SharePoint", ok: false, error: `HTTP ${rBroad.status} or non-JSON`, data: rBroad.json ?? { raw: rBroad.text } };
    }

    const broadAnswer = String(rBroad.json.answer || "");
    if (!broadAnswer || isNoMatchText(broadAnswer)) {
      return {
        source: "SharePoint",
        ok: false,
        error: "NO_MATCH: SharePoint assistant could not find seeded docs or relevant content (likely filename/permissions/indexing).",
        data: { seededAttempt: rSeed.json, broadAttempt: rBroad.json }
      };
    }

    return { source: "SharePoint", ok: true, error: null, data: { ...rBroad.json, note: "Used broad fallback." } };
  }

  // Seeded worked
  return { source: "SharePoint", ok: true, error: null, data: { ...rSeed.json, note: "Used seeded-first lookup." } };
}

/* ----------------------------- 4) Canonical Executive Answer ----------------------------- */

function computeRiskLevel({ sn, sf, sp }) {
  // Forced clarity heuristic (simple + explainable)
  // - If ServiceNow P1 is high => High risk
  // - If at-risk revenue > 250k => High or Medium (depending on ops)
  // - If SharePoint missing => increases uncertainty (but not automatically "High")
  const snData = sn?.data || {};
  const byP = Array.isArray(snData.byPriority) ? snData.byPriority : [];
  const p1 = safeNumber(byP.find((x) => String(x.priority) === "1")?.count, 0);
  const totalHP = safeNumber(snData.totalHighPriority, 0);

  const sfData = sf?.data || {};
  const rev = safeNumber(sfData.atRiskSummary?.totalAmount, 0);

  if (!sn?.ok) return "High — operational visibility is degraded right now.";
  if (p1 >= 50 || totalHP >= 75) return "High — customer experience and SLA commitments are at risk within the next 24 hours.";
  if (rev >= 250000 && (p1 >= 20 || totalHP >= 50)) return "High — combined ops disruption + revenue exposure needs executive coordination.";
  if (rev >= 250000) return "Medium — revenue exposure is material; protect top deals and reduce churn risk.";
  if (!sp?.ok) return "Medium — decision risk: knowledge signals are unavailable; validate impacts with ops owners.";
  return "Low/Medium — monitor closely; no immediate escalation signals beyond baseline.";
}

function buildCanonicalExecutiveAnswer({ question, sources }) {
  const sn = sources.serviceNow;
  const sf = sources.salesforce;
  const sp = sources.sharePoint;

  const snData = sn?.data || {};
  const byP = Array.isArray(snData.byPriority) ? snData.byPriority : [];
  const p1 = safeNumber(byP.find((x) => String(x.priority) === "1")?.count, 0);
  const p2 = safeNumber(byP.find((x) => String(x.priority) === "2")?.count, 0);
  const totalHP = safeNumber(snData.totalHighPriority, 0);

  const sfData = sf?.data || {};
  const acctName = sfData.ebcAccount?.name || "a key account";
  const acctIndustry = sfData.ebcAccount?.industry || "—";
  const revCount = safeNumber(sfData.atRiskSummary?.opportunityCount, 0);
  const revAmount = safeNumber(sfData.atRiskSummary?.totalAmount, 0);

  const spData = sp?.data || {};
  const spAnswerLine = sp?.ok && spData?.answer ? String(spData.answer).trim() : "";
  const spSignal = sp?.ok
    ? (spAnswerLine ? spAnswerLine.split("\n")[0] : "No additional knowledge signals returned.")
    : "Knowledge signal unavailable — SharePoint lookup failed (likely filename/permissions/indexing).";

  // 1) What’s happening (no jargon)
  const whatsHappening = sn?.ok
    ? `There are ${totalHP} high-priority incidents open today (P1 ${p1}, P2 ${p2}), indicating elevated service disruption risk.`
    : `Operational data is partially unavailable because ServiceNow summary failed: ${sn?.error || "unknown error"}.`;

  // 2) Why this matters (business impact)
  const whyMatters = (() => {
    const parts = [];
    if (sn?.ok) parts.push("If unresolved, outages may extend customer downtime and breach SLAs.");
    if (sf?.ok && revCount > 0) parts.push(`We also have ${revCount} at-risk deal(s) worth ~${money(revAmount)} that could slip without proactive coverage.`);
    if (!sp?.ok) parts.push("Additionally, leadership context from SharePoint is missing, increasing decision uncertainty.");
    return parts.join(" ");
  })();

  // 3) Who is impacted (be specific, but don’t hallucinate)
  const whoImpacted = (() => {
    const parts = [];
    if (sf?.ok) parts.push(`Commercial focus: ${acctName} (${acctIndustry}).`);
    if (sn?.ok) parts.push("Operational focus: customers tied to P1/P2 incidents and active SLA commitments.");
    if (sp?.ok && spSignal) parts.push(`Knowledge signals: ${spSignal}`);
    return parts.join(" ");
  })();

  // 4) Risk level (forced clarity)
  const risk = computeRiskLevel({ sn, sf, sp });

  // 5) Leadership attention required (decision-oriented)
  const leadershipAction = (() => {
    const actions = [];
    actions.push("Confirm today’s P1 owners + ETA, and demand a 24-hour stabilization plan (root cause + containment).");
    if (sf?.ok && revCount > 0) actions.push(`Protect revenue: executive sponsor outreach for ${acctName} and top at-risk deals today.`);
    if (!sp?.ok) actions.push("Fix knowledge visibility: validate SharePoint permissions + exact file names for seeded docs and re-test search.");
    return actions.join(" ");
  })();

  const trace = `Traceability: Salesforce(${sf?.ok ? "OK" : "error"}) | ServiceNow(${sn?.ok ? "OK" : "error"}) | SharePoint(${sp?.ok ? "OK" : "error"})`;

  return [
    "EXECUTIVE BRIEF — Today’s Primary Risk & Customer Impact",
    "",
    "1. What’s happening (1 sentence, no jargon)",
    whatsHappening,
    "",
    "2. Why this matters (business impact)",
    whyMatters,
    "",
    "3. Who is impacted (be specific, not generic)",
    whoImpacted,
    "",
    "4. Risk level (explicit, forced clarity)",
    `Risk: ${risk}`,
    "",
    "5. Leadership attention required (decision-oriented)",
    leadershipAction,
    "",
    trace
  ].join("\n");
}

/* ----------------------------- 5) Gemini (Optional) ----------------------------- */

function isGeminiRateLimit(errText = "") {
  const t = String(errText || "").toLowerCase();
  return t.includes("429") || t.includes("resource_exhausted") || t.includes("quota") || t.includes("rate");
}

async function callGeminiPolish({ question, sources }) {
  const key = process.env.GEMINI_API_KEY;
  if (!key) return { used: false, error: "GEMINI_API_KEY not configured." };

  const model = (process.env.GEMINI_MODEL || "gemini-2.5-flash").trim();

  // IMPORTANT: Gemini should "polish" but never shorten/omit sections.
  const instruction =
    `Rewrite the following executive brief in crisp leadership language.\n` +
    `Rules:\n` +
    `- Keep ALL 5 numbered sections and headings exactly.\n` +
    `- Do NOT remove facts.\n` +
    `- Do NOT add invented customers/regions.\n` +
    `- Keep it within 10-14 lines total.\n`;

  const base = buildCanonicalExecutiveAnswer({ question, sources });
  const prompt = `${instruction}\n\nBASE BRIEF:\n${base}`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(key)}`;

  async function attemptOnce() {
    const body = {
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.2, maxOutputTokens: 600 }
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
      getSharePointSummary(question)
    ]);

    const sources = { serviceNow, salesforce, sharePoint };

    // Always build canonical deterministic brief
    let combinedAnswer = buildCanonicalExecutiveAnswer({ question, sources });

    // Gemini optional polish (never required)
    const gemini = await callGeminiPolish({ question, sources });

    if (gemini?.used && gemini?.text) {
      combinedAnswer = gemini.text; // safe because prompt forces keeping all 5 sections
    }

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
