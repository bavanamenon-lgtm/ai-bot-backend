// api/txi-dashboard.js
// Schindler TXI POC - Leadership console aggregator
// - Pulls: ServiceNow incident summary (basic auth), Salesforce strategic account + at-risk opps (jsforce), SharePoint assistant summary (HTTP)
// - Produces: formatted executive answer + per-system trace
// - Hardens: retry + fallback for Gemini (429/503), safe JSON parsing for SharePoint, fallback queries for missing SF custom fields.

import jsforce from "jsforce";

const {
  // Gemini
  GEMINI_API_KEY,
  GEMINI_MODEL, // optional: "gemini-2.0-flash" (default) or "gemini-2.5-flash" etc.

  // ServiceNow
  SN_TXI_URL,     // e.g. https://ven06080.service-now.com/api/dtp/schindler_txi/incident_summary
  SN_USERNAME,
  SN_PASSWORD,

  // Salesforce
  SF_USERNAME,
  SF_PASSWORD,
  SF_TOKEN,
  SF_LOGIN_URL, // e.g. https://login.salesforce.com

  // SharePoint
  SP_CHAT_URL, // IMPORTANT: Full URL to your SharePoint summariser endpoint, e.g. https://ai-bot-backend-black.vercel.app/api/chat-sp
} = process.env;

// -------------------- helpers --------------------

function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*"); // POC only
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
}

function safeJsonParse(text) {
  try {
    return { ok: true, data: JSON.parse(text) };
  } catch {
    return { ok: false, data: null };
  }
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

async function fetchWithTimeout(url, options = {}, timeoutMs = 20000) {
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const resp = await fetch(url, { ...options, signal: controller.signal });
    return resp;
  } finally {
    clearTimeout(t);
  }
}

// -------------------- ServiceNow --------------------

async function getServiceNowSummary() {
  if (!SN_TXI_URL || !SN_USERNAME || !SN_PASSWORD) {
    return {
      source: "ServiceNow",
      error: "ServiceNow env vars not configured (SN_TXI_URL / SN_USERNAME / SN_PASSWORD).",
    };
  }

  const basic = Buffer.from(`${SN_USERNAME}:${SN_PASSWORD}`).toString("base64");

  const resp = await fetchWithTimeout(
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

  const raw = await resp.text();
  const parsed = safeJsonParse(raw);

  if (!resp.ok) {
    return {
      source: "ServiceNow",
      error: `ServiceNow API error (${resp.status})`,
      raw: raw?.slice(0, 500),
    };
  }

  if (!parsed.ok) {
    return {
      source: "ServiceNow",
      error: "ServiceNow returned non-JSON",
      raw: raw?.slice(0, 500),
    };
  }

  // Expecting your endpoint returns something like:
  // { totalHighPriority, byPriority: [{priority,count}], ebcIncidents: [] ... }
  return {
    source: "ServiceNow",
    ...parsed.data,
  };
}

// -------------------- Salesforce --------------------

async function getSalesforceData() {
  // Guard
  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN || !SF_LOGIN_URL) {
    return {
      source: "Salesforce",
      error: "Salesforce env vars not configured (SF_USERNAME / SF_PASSWORD / SF_TOKEN / SF_LOGIN_URL).",
    };
  }

  const conn = new jsforce.Connection({ loginUrl: SF_LOGIN_URL });
  await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);

  // 1) Try find "strategic account" - DO NOT hard depend on Is_EBC_Account__c
  // We'll try in order:
  //   A) Is_EBC_Account__c = true (if exists)
  //   B) Name = 'EBC HQ' (your sample)
  //   C) Rating = 'Hot'
  let ebcAccount = null;
  let accountErr = null;

  const accountQueries = [
    "SELECT Id, Name, Industry, Rating FROM Account WHERE Is_EBC_Account__c = true LIMIT 1",
    "SELECT Id, Name, Industry, Rating FROM Account WHERE Name = 'EBC HQ' LIMIT 1",
    "SELECT Id, Name, Industry, Rating FROM Account WHERE Rating = 'Hot' LIMIT 1",
  ];

  for (const q of accountQueries) {
    try {
      const r = await conn.query(q);
      if (r?.records?.length) {
        const a = r.records[0];
        ebcAccount = {
          id: a.Id,
          name: a.Name,
          industry: a.Industry || null,
          rating: a.Rating || null,
        };
        accountErr = null;
        break;
      }
    } catch (e) {
      accountErr = String(e?.message || e);
      // continue to next fallback
    }
  }

  // 2) At-risk opportunities:
  // Prefer your custom flag Risk_Flag__c. If missing, fallback to Probability < 40 + CloseDate in next 30 days.
  let atRiskOpportunities = [];
  let oppErr = null;

  const oppQueries = [
    "SELECT Id, Name, Amount, StageName, CloseDate, Probability FROM Opportunity WHERE Risk_Flag__c = true ORDER BY CloseDate ASC LIMIT 10",
    "SELECT Id, Name, Amount, StageName, CloseDate, Probability FROM Opportunity WHERE Probability < 40 AND CloseDate = NEXT_N_DAYS:30 ORDER BY CloseDate ASC LIMIT 10",
  ];

  for (const q of oppQueries) {
    try {
      const r = await conn.query(q);
      if (r?.records?.length) {
        atRiskOpportunities = r.records.map((o) => ({
          id: o.Id,
          name: o.Name,
          amount: o.Amount || 0,
          stage: o.StageName || null,
          closeDate: o.CloseDate || null,
          probability: o.Probability ?? null,
        }));
        oppErr = null;
        break;
      }
    } catch (e) {
      oppErr = String(e?.message || e);
    }
  }

  const atRiskSummary = {
    opportunityCount: atRiskOpportunities.length,
    totalAmount: atRiskOpportunities.reduce((s, o) => s + (Number(o.amount) || 0), 0),
  };

  // Important: return BOTH errors separately so TXI can still talk even if one query fails
  const result = {
    source: "Salesforce",
    ebcAccount,
    atRiskSummary,
    atRiskOpportunities,
  };

  if (!ebcAccount && accountErr) result.accountError = accountErr;
  if (!atRiskOpportunities.length && oppErr) result.opportunityError = oppErr;

  return result;
}

// -------------------- SharePoint --------------------

async function getSharePointAnswer(question) {
  if (!SP_CHAT_URL) {
    return {
      source: "SharePoint",
      error: "SP_CHAT_URL env var not configured.",
    };
  }

  // Your chat-sp expects POST { question }
  const resp = await fetchWithTimeout(
    SP_CHAT_URL,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ question }),
    },
    25000
  );

  const raw = await resp.text();
  const parsed = safeJsonParse(raw);

  if (!resp.ok) {
    return {
      source: "SharePoint",
      error: `SharePoint /api/chat-sp error (${resp.status})`,
      raw: raw?.slice(0, 500),
    };
  }

  // Many times your SP endpoint might return HTML (proxy error) → non-JSON
  if (!parsed.ok) {
    return {
      source: "SharePoint",
      error: "SharePoint /api/chat-sp returned non-JSON",
      raw: raw?.slice(0, 500),
    };
  }

  return {
    source: "SharePoint",
    ...parsed.data,
  };
}

// -------------------- Gemini (retry + fallback) --------------------

async function callGeminiExecutive(question, sources) {
  // If no key, skip
  if (!GEMINI_API_KEY) {
    return { used: false, error: "GEMINI_API_KEY not configured." };
  }

  const model = (GEMINI_MODEL || "gemini-2.0-flash").trim();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GEMINI_API_KEY}`;

  // Keep payload small + safe
  const payload = {
    question,
    sources: {
      serviceNow: {
        source: sources?.serviceNow?.source,
        totalHighPriority: sources?.serviceNow?.totalHighPriority,
        byPriority: sources?.serviceNow?.byPriority,
        error: sources?.serviceNow?.error,
      },
      salesforce: {
        source: sources?.salesforce?.source,
        ebcAccount: sources?.salesforce?.ebcAccount,
        atRiskSummary: sources?.salesforce?.atRiskSummary,
        atRiskOpportunities: (sources?.salesforce?.atRiskOpportunities || []).slice(0, 5),
        accountError: sources?.salesforce?.accountError,
        opportunityError: sources?.salesforce?.opportunityError,
      },
      sharePoint: {
        source: sources?.sharePoint?.source,
        answer: sources?.sharePoint?.answer,
        usedFiles: sources?.sharePoint?.usedFiles,
        error: sources?.sharePoint?.error,
      },
    },
  };

  const prompt = `
You are the "Schindler Total Intelligence – Leadership Console".

Write a crisp leadership answer based ONLY on the JSON inputs.
No fluff. No invented facts.

Output format EXACTLY:

**Leadership view (today):**
- 3 bullets max.

**Top 3 issues + business impact:**
1) <Issue name> (Source: <System>)
   - Impact: <1-2 lines, business language>
   - Next move: <one clear action>
2) ...
3) ...

**System status:**
- Salesforce: OK/error
- ServiceNow: OK/error
- SharePoint: OK/error

Rules:
- If a system has an error or no data, say it plainly and use it as a "visibility risk".
- If Salesforce has at-risk opps, include total amount + close dates if present.
- If ServiceNow has high priority counts, include the numbers.
- If SharePoint has usedFiles or answer, summarise in one line.
`.trim();

  const body = {
    contents: [
      {
        parts: [
          { text: prompt },
          { text: "INPUT_JSON:\n" + JSON.stringify(payload) },
        ],
      },
    ],
  };

  // Retry on 429/503 (the exact problem you hit)
  const maxAttempts = 3;
  let lastErr = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const resp = await fetchWithTimeout(
      url,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      },
      25000
    );

    const raw = await resp.text();
    const parsed = safeJsonParse(raw);

    if (resp.ok && parsed.ok) {
      const cand = parsed.data?.candidates?.[0];
      const text = cand?.content?.parts?.[0]?.text;
      if (text) return { used: true, text };
      lastErr = "Gemini returned empty text";
      break;
    }

    // If 429/503, retry with backoff. Otherwise stop.
    if (resp.status === 429 || resp.status === 503) {
      lastErr = `Gemini error (${resp.status})`;
      await sleep(400 * attempt);
      continue;
    }

    lastErr = parsed.ok ? JSON.stringify(parsed.data?.error || parsed.data).slice(0, 300) : raw.slice(0, 300);
    break;
  }

  return { used: false, error: lastErr || "Gemini failed" };
}

// -------------------- Fallback formatter (no Gemini) --------------------

function formatFallbackAnswer(question, sources) {
  const sf = sources.salesforce || {};
  const sn = sources.serviceNow || {};
  const sp = sources.sharePoint || {};

  const sfStatus = sf.error || sf.accountError || sf.opportunityError ? "error" : "OK";
  const snStatus = sn.error ? "error" : "OK";
  const spStatus = sp.error ? "error" : "OK";

  // Build 3 issues based on available data
  const issues = [];

  // 1) Salesforce
  if (sfStatus === "OK" && (sf.atRiskSummary?.opportunityCount > 0)) {
    issues.push({
      title: "Revenue risk on strategic account",
      source: "Salesforce",
      impact: `At-risk opportunities: ${sf.atRiskSummary.opportunityCount} deal(s), total ~$${sf.atRiskSummary.totalAmount}.`,
      action: "Ask the account owner for a save-plan + exec sponsor call today.",
    });
  } else if (sfStatus === "OK") {
    issues.push({
      title: "Sales risk visibility gap",
      source: "Salesforce",
      impact: "No at-risk opps surfaced in current POC data — could be true, or the risk flag logic/data capture is incomplete.",
      action: "Confirm the risk-flag field is used consistently + add 2–3 at-risk sample opps for demo.",
    });
  } else {
    issues.push({
      title: "Salesforce revenue visibility failure",
      source: "Salesforce",
      impact: "Salesforce query failed, so leadership is flying blind on pipeline / customer risk.",
      action: "Fix the SOQL field mapping (custom fields) and permissions; re-run.",
    });
  }

  // 2) ServiceNow
  if (snStatus === "OK" && typeof sn.totalHighPriority === "number") {
    const p1 = (sn.byPriority || []).find((x) => String(x.priority) === "1")?.count ?? "n/a";
    const p2 = (sn.byPriority || []).find((x) => String(x.priority) === "2")?.count ?? "n/a";
    issues.push({
      title: "High-priority incident load",
      source: "ServiceNow",
      impact: `High-priority open: ${sn.totalHighPriority}. (P1: ${p1}, P2: ${p2})`,
      action: "Demand a 24-hour stabilization plan: owners, root causes, and ETA to restore service health.",
    });
  } else {
    issues.push({
      title: "IT visibility gap",
      source: "ServiceNow",
      impact: "ServiceNow data could not be retrieved, so outage/incident risk may be hidden.",
      action: "Fix SN basic-auth/env vars and confirm the endpoint returns JSON.",
    });
  }

  // 3) SharePoint
  if (spStatus === "OK") {
    const usedFiles = sp.usedFiles || [];
    if (usedFiles.length) {
      issues.push({
        title: "Execution/knowledge signal from SharePoint",
        source: "SharePoint",
        impact: `Found ${usedFiles.length} relevant doc(s); summary available.`,
        action: "Use the doc summary to validate delivery risks, deadlines, and owners.",
      });
    } else {
      issues.push({
        title: "Knowledge lookup didn’t match",
        source: "SharePoint",
        impact: "No matching docs found for the question phrasing in the scoped library.",
        action: "Ask using an exact file name (e.g., 'TXI Executive Pack.txt') or a strong keyword from inside the doc.",
      });
    }
  } else {
    issues.push({
      title: "SharePoint knowledge access failure",
      source: "SharePoint",
      impact: "SharePoint summariser errored, so leadership can’t pull execution context from docs.",
      action: "Fix SP_CHAT_URL + confirm Graph permissions + verify the target site/library scope.",
    });
  }

  const lines = [];
  lines.push("**Leadership view (today):**");
  lines.push("- One view across revenue risk, IT stability, and execution knowledge — with traceability by system.");
  lines.push("- When any source fails, that itself becomes a leadership risk (visibility blind spot).");
  lines.push("");
  lines.push(`**Question asked:** ${question}`);
  lines.push("");
  lines.push("**Top 3 issues + business impact:**");
  issues.slice(0, 3).forEach((it, i) => {
    lines.push(`${i + 1}) **${it.title}** *(Source: ${it.source})*`);
    lines.push(`- Impact: ${it.impact}`);
    lines.push(`- Next move: ${it.action}`);
    lines.push("");
  });
  lines.push("**System status:**");
  lines.push(`- Salesforce: ${sfStatus}`);
  lines.push(`- ServiceNow: ${snStatus}`);
  lines.push(`- SharePoint: ${spStatus}`);

  return lines.join("\n");
}

// -------------------- Handler --------------------

export default async function handler(req, res) {
  setCors(res);

  if (req.method === "OPTIONS") {
    res.status(200).end();
    return;
  }

  if (req.method !== "POST") {
    res.status(405).json({ error: 'Use POST with JSON body { "question": "..." }' });
    return;
  }

  try {
    const { question } = req.body || {};
    if (!question || typeof question !== "string") {
      res.status(400).json({ error: 'Missing "question" in request body.' });
      return;
    }

    // Pull from systems in parallel (faster + less timeouts)
    const [serviceNow, salesforce, sharePoint] = await Promise.allSettled([
      getServiceNowSummary(),
      getSalesforceData(),
      getSharePointAnswer(question),
    ]);

    const sources = {
      serviceNow: serviceNow.status === "fulfilled" ? serviceNow.value : { source: "ServiceNow", error: String(serviceNow.reason) },
      salesforce: salesforce.status === "fulfilled" ? salesforce.value : { source: "Salesforce", error: String(salesforce.reason) },
      sharePoint: sharePoint.status === "fulfilled" ? sharePoint.value : { source: "SharePoint", error: String(sharePoint.reason) },
    };

    // Gemini with retry/fallback
    const gemini = await callGeminiExecutive(question, sources);

    const combinedAnswer = gemini.used
      ? gemini.text
      : formatFallbackAnswer(question, sources);

    res.status(200).json({
      question,
      combinedAnswer,
      sources,
      generatedAt: new Date().toISOString(),
      gemini: gemini.used ? { used: true, model: (GEMINI_MODEL || "gemini-2.0-flash") } : { used: false, error: gemini.error },
    });
  } catch (err) {
    res.status(500).json({
      error: "Internal error in /api/txi-dashboard",
      details: String(err?.message || err),
    });
  }
}
