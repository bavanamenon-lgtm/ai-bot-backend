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
  return (
    errText.includes("429") ||
    errText.toLowerCase().includes("resource_exhausted") ||
    errText.toLowerCase().includes("quota") ||
    errText.toLowerCase().includes("rate")
  );
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

    // Try JSON parse; if it fails, return raw text
    let json = null;
    try {
      json = JSON.parse(text);
    } catch {
      json = null;
    }

    return {
      ok: resp.ok,
      status: resp.status,
      text,
      json,
    };
  } finally {
    clearTimeout(id);
  }
}

/** -----------------------------
 *  1) ServiceNow fetch
 *  ----------------------------- */
async function getServiceNowSummary() {
  const SN_TXI_URL = process.env.SN_TXI_URL; // e.g. https://ven06080.service-now.com/api/dtp/schindler_txi/incident_summary
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
    return {
      source: "ServiceNow",
      error: `ServiceNow HTTP ${r.status}`,
      raw: r.json ?? r.text,
    };
  }

  // Your endpoint already returns a good JSON structure.
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

  // Strategy:
  // A) Try to locate an EBC-ish account without relying on custom fields.
  //    Use the account you created: "EBC HQ".
  // B) Pull at-risk opps using common checkbox API names (At_Risk__c, Risk_Flag__c).
  //    If those fields don't exist, fallback to stage/probability based shortlist.

  let ebcAccount = null;

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
  } catch (e) {
    // Keep going; we’ll fallback later.
  }

  // If not found, pick a "Hot" account as fallback.
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

  // Try at-risk opps using custom fields (your UI shows Risk Flag + At Risk)
  let oppQueryTried = [];
  const oppQueries = [
    // most likely custom fields
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
    // fallback: treat low probability + near close date as risk
    `SELECT Id, Name, Amount, StageName, CloseDate, Probability
     FROM Opportunity
     WHERE AccountId = '${ebcAccount.id}'
     ORDER BY CloseDate ASC
     LIMIT 10`,
  ];

  let oppRecords = null;
  let oppError = null;

  for (const q of oppQueries) {
    oppQueryTried.push(q);
    try {
      const o = await conn.query(q);
      oppRecords = o?.records || [];
      oppError = null;
      break;
    } catch (e) {
      oppError = e;
      continue;
    }
  }

  if (!oppRecords) {
    return {
      source: "Salesforce",
      ebcAccount,
      error: `Opportunity query failed: ${oppError?.message || String(oppError)}`,
      debug: { oppQueryTried },
    };
  }

  // Normalize
  const atRiskOpportunities = oppRecords.map((r) => ({
    id: r.Id,
    name: r.Name,
    amount: safeNumber(r.Amount, 0),
    stage: r.StageName,
    closeDate: r.CloseDate,
    probability: safeNumber(r.Probability, 0),
  }));

  // If we used fallback query, we should *compute* "at risk" as probability <= 30
  const computedAtRisk = atRiskOpportunities.filter((o) => o.probability <= 30);

  const useList = atRiskOpportunities.some((o) => o.name) ? atRiskOpportunities : [];
  const listForSummary =
    // if At_Risk__c/Risk_Flag__c worked, take those
    (oppQueryTried[0] && oppQueryTried[0].includes("At_Risk__c") && useList.length ? useList : null) ||
    (oppQueryTried[1] && oppQueryTried[1].includes("Risk_Flag__c") && useList.length ? useList : null) ||
    // else computed risk list
    (computedAtRisk.length ? computedAtRisk : useList);

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
  const SP_CHAT_URL = process.env.SP_CHAT_URL; // your working SharePoint assistant endpoint
  if (!SP_CHAT_URL) {
    return { source: "SharePoint", error: "SP_CHAT_URL env var not configured." };
  }

  // First attempt: ask the question as-is
  const r1 = await fetchJson(
    SP_CHAT_URL,
    {
      method: "POST",
      headers: JSON_HEADERS,
      body: JSON.stringify({ question }),
    },
    25000
  );

  if (r1.ok && r1.json) {
    // if it says "no matching files", do a second attempt with seeded filenames/keywords
    const ans = String(r1.json.answer || "");
    const noMatch = ans.toLowerCase().includes("couldn't find") || ans.toLowerCase().includes("no matching");

    if (!noMatch) return { source: "SharePoint", ...r1.json };

    // Second attempt: force it to look for your seeded files
    const seededPrompt =
      `Search specifically for these files and summarise what’s relevant for leadership risks:\n` +
      `1) Annual EBC Review Notes.txt\n` +
      `2) EBC_Account_Health_Risk.docx\n` +
      `3) IT_Operations_Weekly_Report.docx\n\n` +
      `If found, return 3 bullets: Risk, Customer impact, Action.\n` +
      `Original question: ${question}`;

    const r2 = await fetchJson(
      SP_CHAT_URL,
      {
        method: "POST",
        headers: JSON_HEADERS,
        body: JSON.stringify({ question: seededPrompt }),
      },
      25000
    );

    if (r2.ok && r2.json) return { source: "SharePoint", ...r2.json };

    return {
      source: "SharePoint",
      error: `SharePoint assistant returned HTTP ${r2.status}`,
      raw: r2.json ?? r2.text,
    };
  }

  return {
    source: "SharePoint",
    error: `SharePoint assistant returned HTTP ${r1.status}`,
    raw: r1.json ?? r1.text,
  };
}

/** -----------------------------
 *  4) Gemini (2.5/3) with fallback
 *  ----------------------------- */
async function callGemini({ question, sources }) {
  const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
  if (!GEMINI_API_KEY) {
    return { used: false, error: "GEMINI_API_KEY not configured." };
  }

  // You can override via env; else we try best → fallback
  const preferred = (process.env.GEMINI_MODEL || "").trim();
  const modelChain = [
    preferred,
    "gemini-2.5-flash",
    "gemini-2.5-pro",
    "gemini-2.0-flash",
  ].filter(Boolean);

  const systemInstruction =
    `You are an executive TXI assistant. Output MUST be short, structured, and action-oriented.\n` +
    `Format EXACTLY like:\n` +
    `TITLE (1 line)\n` +
    `1) Risk headline — So what (1 sentence) — Do now (1 sentence)\n` +
    `2) ...\n` +
    `3) ...\n` +
    `Then a final line: "Traceability: Salesforce | ServiceNow | SharePoint"\n` +
    `No long paragraphs. No filler.\n`;

  const prompt =
    `${systemInstruction}\n` +
    `Question: ${question}\n\n` +
    `Data (ground truth JSON). Do not invent fields not present:\n` +
    JSON.stringify(sources, null, 2);

  // v1beta generateContent endpoint
  async function tryModel(model) {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(
      GEMINI_API_KEY
    )}`;

    const body = {
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: {
        temperature: 0.2,
        maxOutputTokens: 450,
      },
    };

    const r = await fetchJson(
      url,
      { method: "POST", headers: JSON_HEADERS, body: JSON.stringify(body) },
      25000
    );

    if (!r.ok) {
      const errText = r.text || "";
      return { ok: false, status: r.status, errText, raw: r.json ?? r.text };
    }

    const text =
      r.json?.candidates?.[0]?.content?.parts?.map((p) => p.text).join("") ||
      r.json?.candidates?.[0]?.content?.parts?.[0]?.text ||
      "";

    if (!text.trim()) {
      return { ok: false, status: 500, errText: "Empty Gemini response", raw: r.json };
    }

    return { ok: true, text };
  }

  // retry logic (short) + model fallback
  let lastErr = null;

  for (const model of modelChain) {
    // 2 attempts per model
    for (let attempt = 1; attempt <= 2; attempt++) {
      const out = await tryModel(model);

      if (out.ok) {
        return { used: true, model, text: out.text };
      }

      lastErr = out;
      // If rate-limited, wait a bit then either retry or move to next model
      if (isGeminiRateLimit(out.errText) || out.status === 429 || out.status === 503) {
        await sleep(900 * attempt);
        continue;
      } else {
        // non-rate error → move to next model
        break;
      }
    }
  }

  return { used: false, error: `Gemini failed. Last: HTTP ${lastErr?.status}`, raw: lastErr?.raw };
}

/** -----------------------------
 *  5) Local executive formatter (no Gemini needed)
 *  ----------------------------- */
function buildExecutiveAnswer({ question, sources }) {
  const sf = sources.salesforce;
  const sn = sources.serviceNow;
  const sp = sources.sharePoint;

  // Revenue risk
  const sfRisk =
    sf?.atRiskSummary
      ? `Revenue risk on ${sf?.ebcAccount?.name || "strategic account"}: ${sf.atRiskSummary.opportunityCount} deal(s), total ${money(sf.atRiskSummary.totalAmount)}.`
      : sf?.error
      ? `Sales risk visibility gap: ${String(sf.error).slice(0, 140)}.`
      : `Sales risk: insufficient data.`;

  // IT risk
  const p1 = sn?.byPriority?.find((x) => String(x.priority) === "1")?.count;
  const p2 = sn?.byPriority?.find((x) => String(x.priority) === "2")?.count;
  const high = sn?.totalHighPriority;

  const snRisk =
    sn?.error
      ? `IT visibility gap: ${String(sn.error).slice(0, 140)}.`
      : `IT stability risk: ${safeNumber(high)} high-priority incidents open (P1 ${safeNumber(p1)}, P2 ${safeNumber(p2)}).`;

  // Knowledge risk
  const spAns = sp?.answer ? String(sp.answer) : "";
  const spRisk =
    sp?.error
      ? `Knowledge visibility gap: ${String(sp.error).slice(0, 140)}.`
      : spAns.toLowerCase().includes("couldn't find")
      ? `Knowledge signal weak: no matching docs found in scope for this phrasing (try exact file names).`
      : `Knowledge signal available: insights pulled from SharePoint.`;

  // Do-now actions (tight)
  const lines = [
    `TXI Leadership Snapshot — today`,
    `1) Revenue — ${sfRisk} Do now: exec sponsor + “save plan” for the top 2 deals today.`,
    `2) Operations — ${snRisk} Do now: 24-hour stabilization plan (owners, root cause, ETA).`,
    `3) Execution knowledge — ${spRisk} Do now: run targeted doc query: “EBC_Account_Health_Risk” + “IT_Operations_Weekly_Report”.`,
    `Traceability: Salesforce | ServiceNow | SharePoint`,
  ];

  return lines.join("\n");
}

/** -----------------------------
 *  Handler
 *  ----------------------------- */
export default async function handler(req, res) {
  allowCors(res);

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: 'Use POST with JSON body { "question": "..." }' });
  }

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
    // Fetch all sources in parallel
    const [serviceNow, salesforce, sharePoint] = await Promise.all([
      getServiceNowSummary(),
      getSalesforceSummary(),
      getSharePointSummary(question),
    ]);

    const sources = { serviceNow, salesforce, sharePoint };

    // Try Gemini (optional), but NEVER block output on it
    const gemini = await callGemini({ question, sources });

    const combinedAnswer =
      gemini?.used && gemini?.text
        ? gemini.text
        : buildExecutiveAnswer({ question, sources });

    return res.status(200).json({
      question,
      combinedAnswer,
      sources,
      gemini: gemini?.used ? { used: true, model: gemini.model } : { used: false, error: gemini?.error || null },
      generatedAt: new Date().toISOString(),
    });
  } catch (e) {
    return res.status(500).json({
      error: "FUNCTION_INVOCATION_FAILED",
      detail: e?.message || String(e),
    });
  }
}
