// api/txi-dashboard.js
// Schindler Total Intelligence – Leadership Console (TXI POC)
// FINAL HARDENED VERSION:
// - Canonical Executive Template (5 sections) ALWAYS
// - Gemini can "polish" ONLY if it preserves all sections (guarded)
// - SharePoint seeded lookup targets Site/Library/Folder explicitly
// - SharePoint "no match" correctly sets ok:false (no fake green pills)
//
// POST /api/txi-dashboard { "question": "..." }
//
// Env vars required:
// SN_TXI_URL, SN_USERNAME, SN_PASSWORD
// SF_USERNAME, SF_PASSWORD, SF_TOKEN, SF_LOGIN_URL(optional)
// SP_CHAT_URL
// Optional: GEMINI_API_KEY, GEMINI_MODEL(optional; default gemini-2.5-flash)

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

function isNoMatchText(s = "") {
  const t = String(s || "").toLowerCase();
  return (
    t.includes("no_match") ||
    t.includes("couldn't find") ||
    t.includes("could not find") ||
    t.includes("no matching") ||
    t.includes("try using an exact file name") ||
    t.includes("documents library") && t.includes("couldn't find")
  );
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

  // Portable heuristic (no custom fields):
  // "at-risk" = low probability OR closing soon
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

/* ----------------------------- 3) SharePoint (seeded + explicit target) ----------------------------- */
async function askSharePointAssistant(prompt) {
  const SP_CHAT_URL = process.env.SP_CHAT_URL;
  if (!SP_CHAT_URL) {
    return { ok: false, status: 0, json: null, text: "Missing SP_CHAT_URL" };
  }
  return fetchJson(
    SP_CHAT_URL,
    { method: "POST", headers: JSON_HEADERS, body: JSON.stringify({ question: prompt }) },
    30000
  );
}

async function getSharePointSummary(question) {
  const SP_CHAT_URL = process.env.SP_CHAT_URL;
  if (!SP_CHAT_URL) {
    return { source: "SharePoint", ok: false, error: "Missing SP_CHAT_URL", data: null };
  }

  // Your screenshot proves the site is "Vation GTM" and library is "Documents"
  // Also likely a Teams-connected folder "General".
  const seededPrompt =
`You are connected to Microsoft SharePoint / Teams.

TARGET LOCATION (must follow exactly):
- Site name: "Vation GTM"
- Library name: "Documents"
- Folder(s) to check: root of Documents AND "General"

Step 1: List the filenames you can see in that library/folder (max 15).
Step 2: Open and extract leadership signals ONLY from these exact files (if present):
- Annual EBC Review Notes.txt
- EBC_Account_Health_Risk.docx
- IT_Operations_Weekly_Report.docx
- Sales_Risk_Accounts_List.docx

Output format MUST be:
Key risk signals:
- Risk | Customer impact | Suggested action
- Risk | Customer impact | Suggested action
- Risk | Customer impact | Suggested action

If you cannot list files OR cannot find these exact filenames, output EXACTLY: NO_MATCH

User question context: ${question}`;

  // Attempt 1: seeded + explicit target
  const rSeed = await askSharePointAssistant(seededPrompt);
  if (!rSeed.ok || !rSeed.json) {
    return { source: "SharePoint", ok: false, error: `HTTP ${rSeed.status} or non-JSON`, data: rSeed.json ?? { raw: rSeed.text } };
  }

  const seededAnswer = String(rSeed.json.answer || "");
  if (!seededAnswer || seededAnswer.trim() === "NO_MATCH" || isNoMatchText(seededAnswer)) {
    // Attempt 2: broad question but still anchored to the location
    const anchoredBroadPrompt =
`Use the same TARGET LOCATION:
- Site: "Vation GTM"
- Library: "Documents"
- Folder(s): root and "General"

Answer the user question using those documents only.
If nothing relevant is found, output EXACTLY: NO_MATCH

User question: ${question}`;

    const rBroad = await askSharePointAssistant(anchoredBroadPrompt);

    if (!rBroad.ok || !rBroad.json) {
      return { source: "SharePoint", ok: false, error: `HTTP ${rBroad.status} or non-JSON`, data: rBroad.json ?? { raw: rBroad.text } };
    }

    const broadAnswer = String(rBroad.json.answer || "");
    if (!broadAnswer || broadAnswer.trim() === "NO_MATCH" || isNoMatchText(broadAnswer)) {
      return {
        source: "SharePoint",
        ok: false,
        error: "NO_MATCH: SharePoint assistant cannot enumerate/find files in 'Vation GTM' → 'Documents' (root/General). Likely wrong drive mapping or permissions for the assistant identity.",
        data: { seededAttempt: rSeed.json, broadAttempt: rBroad.json }
      };
    }

    return { source: "SharePoint", ok: true, error: null, data: { ...rBroad.json, note: "Used anchored broad fallback." } };
  }

  return { source: "SharePoint", ok: true, error: null, data: { ...rSeed.json, note: "Used seeded explicit lookup." } };
}

/* ----------------------------- 4) Canonical Executive Template ----------------------------- */
function computeRiskLevel({ sn, sf, sp }) {
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
  if (!sp?.ok) return "Medium — leadership context from SharePoint is missing; validate impacts with ops owners.";
  return "Low/Medium — monitor closely; no immediate escalation beyond baseline.";
}

function buildCanonicalExecutiveAnswer({ sources }) {
  const sn = sources.serviceNow || {};
  const sf = sources.salesforce || {};
  const sp = sources.sharePoint || {};

  const snData = sn.data || {};
  const byP = Array.isArray(snData.byPriority) ? snData.byPriority : [];
  const p1 = safeNumber(byP.find((x) => String(x.priority) === "1")?.count, 0);
  const p2 = safeNumber(byP.find((x) => String(x.priority) === "2")?.count, 0);
  const totalHP = safeNumber(snData.totalHighPriority, 0);

  const sfData = sf.data || {};
  const acctName = sfData.ebcAccount?.name || "a key account";
  const acctIndustry = sfData.ebcAccount?.industry || "—";
  const revCount = safeNumber(sfData.atRiskSummary?.opportunityCount, 0);
  const revAmount = safeNumber(sfData.atRiskSummary?.totalAmount, 0);

  const spData = sp.data || {};
  const spAnswerLine = sp.ok && spData?.answer ? String(spData.answer).trim() : "";
  const spSignal = sp.ok
    ? (spAnswerLine ? spAnswerLine.split("\n")[0] : "No additional SharePoint signal returned.")
    : "SharePoint leadership context unavailable (assistant could not locate documents in Vation GTM → Documents).";

  const whatsHappening = sn.ok
    ? `A spike of high-priority incidents is visible today: ${totalHP} open (P1 ${p1}, P2 ${p2}).`
    : `ServiceNow signal is unavailable right now: ${sn.error || "unknown error"}.`;

  const whyMatters = (() => {
    const parts = [];
    if (sn.ok) parts.push("If unresolved, this can extend customer downtime and risk SLA breaches.");
    if (sf.ok && revCount > 0) parts.push(`We also have ${revCount} at-risk deal(s) worth ~${money(revAmount)} requiring proactive coverage.`);
    if (!sp.ok) parts.push("SharePoint context is missing, increasing decision uncertainty for customer/region impact.");
    return parts.join(" ");
  })();

  const whoImpacted = (() => {
    const parts = [];
    if (sf.ok) parts.push(`Commercial focus: ${acctName} (${acctIndustry}).`);
    if (sn.ok) parts.push("Operational focus: customers tied to P1/P2 incidents and active SLA commitments.");
    if (sp.ok && spSignal) parts.push(`Document signal: ${spSignal}`);
    return parts.join(" ");
  })();

  const risk = computeRiskLevel({ sn, sf, sp });

  const leadershipAction = (() => {
    const actions = [];
    actions.push("Confirm P1 owners + ETA now, and demand a 24-hour stabilization plan (containment + root cause).");
    if (sf.ok && revCount > 0) actions.push(`Protect revenue: executive sponsor outreach for ${acctName} and top at-risk deals today.`);
    if (!sp.ok) actions.push("Fix SharePoint: validate assistant permissions + correct drive mapping to Vation GTM → Documents (root/General), then re-test seeded filenames.");
    return actions.join(" ");
  })();

  const trace = `Traceability: Salesforce(${sf.ok ? "OK" : "error"}) | ServiceNow(${sn.ok ? "OK" : "error"}) | SharePoint(${sp.ok ? "OK" : "error"})`;

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

/* ----------------------------- 5) Gemini (optional polish with hard guard) ----------------------------- */
function looksLikeCanonicalTemplate(text) {
  const t = String(text || "");
  // Must contain the 5 sections + Traceability (guard against tiny outputs)
  const must = [
    "EXECUTIVE BRIEF",
    "1. What’s happening",
    "2. Why this matters",
    "3. Who is impacted",
    "4. Risk level",
    "5. Leadership attention required",
    "Traceability:"
  ];
  return must.every((m) => t.includes(m)) && t.length >= 350; // hard minimum to prevent 80-char junk
}

function isGeminiRateLimit(errText = "") {
  const t = String(errText || "").toLowerCase();
  return t.includes("429") || t.includes("resource_exhausted") || t.includes("quota") || t.includes("rate");
}

async function callGeminiPolish(canonicalBrief) {
  const key = process.env.GEMINI_API_KEY;
  if (!key) return { used: false, error: "GEMINI_API_KEY not configured." };

  const model = (process.env.GEMINI_MODEL || "gemini-2.5-flash").trim();

  const instruction =
`Rewrite the executive brief in crisp leadership language.

ABSOLUTE RULES:
- Keep ALL headings and all 5 sections exactly as-is.
- Keep the "Traceability:" line.
- Do NOT remove facts.
- Do NOT add invented customers/regions.
- Do NOT shorten into a one-liner.
- Keep it readable and executive-grade.`;

  const prompt = `${instruction}\n\nBRIEF:\n${canonicalBrief}`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(key)}`;

  async function attemptOnce() {
    const body = {
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.2, maxOutputTokens: 700 }
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

    // 1) Always generate canonical deterministic brief
    const canonical = buildCanonicalExecutiveAnswer({ sources });

    // 2) Gemini polish is optional; accept ONLY if it preserves template + length
    let combinedAnswer = canonical;
    const gemini = await callGeminiPolish(canonical);

    if (gemini?.used && gemini?.text && looksLikeCanonicalTemplate(gemini.text)) {
      combinedAnswer = gemini.text;
    } else {
      // If Gemini produced junk/short output, ignore it (never override canonical)
      // Keep gemini metadata for debug transparency
    }

    return res.status(200).json({
      question,
      combinedAnswer,
      sources,
      gemini: (gemini?.used && gemini?.text && looksLikeCanonicalTemplate(gemini.text))
        ? { used: true, model: gemini.model }
        : { used: false, error: gemini?.error || "Gemini output rejected by template guard." },
      generatedAt: new Date().toISOString()
    });
  } catch (e) {
    return res.status(500).json({ error: "FUNCTION_INVOCATION_FAILED", detail: e?.message || String(e) });
  }
}
