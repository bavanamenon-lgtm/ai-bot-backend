// api/txi-dashboard.js
// TXI Dashboard Aggregator (ServiceNow + Salesforce + SharePoint) + Gemini summary (optional)
//
// Goals:
// - Never fail the CEO dashboard: always return a well-formatted answer even if Gemini fails.
// - Make SharePoint deterministic: search via TXI keywords (not the full CEO question).
// - Make Salesforce robust: auto-detect custom fields via describe() so no INVALID_FIELD surprises.
// - Gemini model configurable + retry + graceful fallback.
//
// Required env vars:
//   GEMINI_API_KEY
//   GEMINI_MODEL (optional) e.g. "gemini-2.0-flash" (default), "gemini-1.5-flash", "gemini-2.5-flash" if available
//
//   SN_TXI_URL (e.g. https://ven06080.service-now.com/api/dtp/schindler_txi/incident_summary)
//   SN_USERNAME
//   SN_PASSWORD
//
//   SF_USERNAME
//   SF_PASSWORD
//   SF_TOKEN
//   SF_LOGIN_URL (e.g. https://login.salesforce.com or sandbox url)
//
//   SP_CHAT_URL (Full URL to your SharePoint endpoint, e.g. https://<your-vercel>.vercel.app/api/chat-sp)
//
// Notes:
// - This endpoint expects POST JSON { "question": "..." }.
// - It returns JSON with { combinedAnswer, sources, gemini }.

import jsforce from "jsforce";

export const config = {
  runtime: "nodejs",
};

function json(res, status, body) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json");
  res.end(JSON.stringify(body, null, 2));
}

function safeString(x) {
  return (x == null ? "" : String(x)).trim();
}

function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
}

function maskError(e) {
  return safeString(e && (e.message || e)).slice(0, 600);
}

// -------------------- ServiceNow --------------------

async function fetchServiceNowSummary() {
  const url = process.env.SN_TXI_URL;
  const u = process.env.SN_USERNAME;
  const p = process.env.SN_PASSWORD;

  if (!url || !u || !p) {
    return {
      source: "ServiceNow",
      error: "ServiceNow env vars not configured (SN_TXI_URL / SN_USERNAME / SN_PASSWORD).",
    };
  }

  const auth = Buffer.from(`${u}:${p}`).toString("base64");

  const resp = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Basic ${auth}`,
      Accept: "application/json",
    },
  });

  const raw = await resp.text();
  let data;
  try {
    data = JSON.parse(raw);
  } catch {
    return {
      source: "ServiceNow",
      error: `ServiceNow returned non-JSON (HTTP ${resp.status}).`,
      raw: raw.slice(0, 500),
    };
  }

  if (!resp.ok) {
    return {
      source: "ServiceNow",
      error: `ServiceNow API error (HTTP ${resp.status}).`,
      raw: JSON.stringify(data).slice(0, 500),
    };
  }

  // Expecting your Scripted REST output already shaped.
  return data;
}

// -------------------- Salesforce --------------------

async function sfLogin() {
  const username = process.env.SF_USERNAME;
  const password = process.env.SF_PASSWORD;
  const token = process.env.SF_TOKEN;
  const loginUrl = process.env.SF_LOGIN_URL || "https://login.salesforce.com";

  if (!username || !password || !token) {
    throw new Error("Salesforce env vars not configured (SF_USERNAME / SF_PASSWORD / SF_TOKEN).");
  }

  const conn = new jsforce.Connection({ loginUrl });
  await conn.login(username, password + token);
  return conn;
}

async function sfFindFirstExistingField(conn, sobjectName, candidates) {
  const desc = await conn.sobject(sobjectName).describe();
  const fieldSet = new Set((desc.fields || []).map((f) => f.name));
  for (const c of candidates) {
    if (fieldSet.has(c)) return c;
  }
  return null;
}

async function fetchSalesforceSummary() {
  try {
    const conn = await sfLogin();

    // 1) Find “EBC account” flag field (robust)
    const ebcField = await sfFindFirstExistingField(conn, "Account", [
      "Is_EBC_Account__c",
      "EBC_Account__c",
      "EBC_Flag__c",
      "IsEBC__c",
    ]);

    // 2) Find “risk flag” field on Opportunity (robust)
    const oppRiskField = await sfFindFirstExistingField(conn, "Opportunity", [
      "Risk_Flag__c",
      "Is_At_Risk__c",
      "At_Risk__c",
      "RiskFlag__c",
    ]);

    // If missing, don’t die — report as visibility gap.
    let ebcAccount = null;
    if (ebcField) {
      const q = `SELECT Id, Name, Industry, Rating FROM Account WHERE ${ebcField} = true LIMIT 1`;
      const r = await conn.query(q);
      ebcAccount = (r.records && r.records[0]) ? {
        id: r.records[0].Id,
        name: r.records[0].Name,
        industry: r.records[0].Industry || null,
        rating: r.records[0].Rating || null,
      } : null;
    }

    let atRiskOpps = [];
    if (oppRiskField) {
      // If you want only the EBC account’s opps, add AccountId filter (optional)
      const q =
        `SELECT Id, Name, Amount, StageName, CloseDate, Probability ` +
        `FROM Opportunity WHERE ${oppRiskField} = true ORDER BY CloseDate ASC LIMIT 10`;
      const r = await conn.query(q);
      atRiskOpps = (r.records || []).map((o) => ({
        id: o.Id,
        name: o.Name,
        amount: o.Amount || 0,
        stage: o.StageName || "",
        closeDate: o.CloseDate || "",
        probability: o.Probability || 0,
      }));
    }

    const totalAmount = atRiskOpps.reduce((sum, o) => sum + (Number(o.amount) || 0), 0);

    return {
      source: "Salesforce",
      ebcAccount,
      atRiskSummary: {
        opportunityCount: atRiskOpps.length,
        totalAmount,
      },
      atRiskOpportunities: atRiskOpps,
      detectedFields: {
        ebcField: ebcField || null,
        oppRiskField: oppRiskField || null,
      },
      warnings: [
        ...(ebcField ? [] : ["EBC account flag field not found on Account (expected one of: Is_EBC_Account__c / EBC_Account__c / etc)."]),
        ...(oppRiskField ? [] : ["Risk flag field not found on Opportunity (expected one of: Risk_Flag__c / Is_At_Risk__c / etc)."]),
      ],
    };
  } catch (e) {
    return {
      source: "Salesforce",
      error: maskError(e),
    };
  }
}

// -------------------- SharePoint --------------------

function buildSharePointQuery(question) {
  // IMPORTANT: don’t pass the entire CEO question as search term.
  // Use deterministic TXI keywords so the 3 files are found.
  const q = safeString(question).toLowerCase();

  // If the question is about “risks / impacts / issues”, we search for these keywords.
  // These match your file names and content strongly.
  const keywords = [
    "EBC_Account_Health_Risk",
    "IT_Operations_Weekly_Report",
    "Sales_Risk_Accounts_List",
    "executive pack",
    "risk",
    "incident trends",
    "at-risk opportunities",
  ];

  // If user explicitly mentions a file name, use it.
  const fileMatch = question && question.match(/([A-Za-z0-9_\- ]+\.(docx|xlsx|txt|csv))/i);
  if (fileMatch && fileMatch[1]) return fileMatch[1];

  // Else, return a compact query that Graph search will match.
  if (q.includes("risk") || q.includes("impact") || q.includes("issues")) {
    return keywords.slice(0, 4).join(" ");
  }

  return "TXI Executive Pack risk incident opportunities";
}

async function fetchSharePointSummary(question) {
  const spUrl = process.env.SP_CHAT_URL;
  if (!spUrl) {
    return { source: "SharePoint", error: "SP_CHAT_URL env var not configured." };
  }

  const spQuestion = buildSharePointQuery(question);

  const resp = await fetch(spUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ question: spQuestion }),
  });

  const raw = await resp.text();
  let data;
  try {
    data = JSON.parse(raw);
  } catch {
    return {
      source: "SharePoint",
      error: "SharePoint /api/chat-sp returned non-JSON",
      raw: raw.slice(0, 600),
    };
  }

  if (!resp.ok) {
    return {
      source: "SharePoint",
      error: `SharePoint /api/chat-sp error (${resp.status})`,
      raw: JSON.stringify(data).slice(0, 600),
    };
  }

  return {
    source: "SharePoint",
    answer: data.answer || "",
    usedFiles: data.usedFiles || [],
    candidateFiles: data.candidateFiles || [],
  };
}

// -------------------- Gemini --------------------

async function callGeminiLeadershipSummary({ question, serviceNow, salesforce, sharePoint }) {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return { used: false, error: "GEMINI_API_KEY not configured." };
  }

  const model = process.env.GEMINI_MODEL || "gemini-2.0-flash";

  const url =
    `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${apiKey}`;

  // Keep payload small to reduce 429 risk.
  const payload = {
    question,
    serviceNow: {
      totalHighPriority: serviceNow?.totalHighPriority ?? null,
      byPriority: serviceNow?.byPriority ?? null,
    },
    salesforce: {
      ebcAccount: salesforce?.ebcAccount ?? null,
      atRiskSummary: salesforce?.atRiskSummary ?? null,
      atRiskOpportunities: salesforce?.atRiskOpportunities ?? [],
      warnings: salesforce?.warnings ?? [],
    },
    sharePoint: {
      usedFiles: (sharePoint?.usedFiles || []).slice(0, 3).map((f) => ({ name: f.name, lastModified: f.lastModified })),
      answer: safeString(sharePoint?.answer).slice(0, 1200),
    },
  };

  const prompt = `
You are writing a CEO-friendly operational intelligence brief.

Rules:
- Output MUST be concise and usable in a leadership meeting.
- Use this exact structure:

**Leadership view (today):**
- <2 bullets max>

**Top 3 issues + business impact:**
1) <Title> *(Source: Salesforce/ServiceNow/SharePoint)*
- Impact: <1 line>
- Next move (today): <1 line>

2) ...
3) ...

**System status:** Salesforce: <OK/error> | ServiceNow: <OK/error> | SharePoint: <OK/error>

- Do NOT mention “POC”, “JSON”, “API”, “env vars”.
- If a source has an error, treat it as “visibility blind spot” and give a practical next move.
- If SharePoint “no docs found”, suggest the exact file names that likely exist.
`.trim();

  async function geminiAttempt() {
    const resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ parts: [{ text: `${prompt}\n\nDATA:\n${JSON.stringify(payload)}` }] }],
      }),
    });

    const raw = await resp.text();
    let data;
    try {
      data = JSON.parse(raw);
    } catch {
      throw new Error(`Gemini non-JSON (HTTP ${resp.status})`);
    }
    if (!resp.ok) {
      const code = data?.error?.code || resp.status;
      const msg = data?.error?.message || "Gemini API error";
      const err = new Error(`Gemini error (${code}): ${msg}`);
      err._code = code;
      throw err;
    }

    const text =
      data?.candidates?.[0]?.content?.parts?.[0]?.text ||
      "";
    return text.trim();
  }

  // Retry once on 429/503
  try {
    const text = await geminiAttempt();
    return { used: true, model, text };
  } catch (e1) {
    const code = e1?._code;
    if (code === 429 || code === 503) {
      await new Promise((r) => setTimeout(r, 650)); // tiny backoff
      try {
        const text = await geminiAttempt();
        return { used: true, model, text };
      } catch (e2) {
        return { used: false, error: maskError(e2) };
      }
    }
    return { used: false, error: maskError(e1) };
  }
}

// -------------------- Fallback formatter (no Gemini) --------------------

function systemOk(obj) {
  return obj && !obj.error;
}

function formatFallbackAnswer(question, sn, sf, sp) {
  const sfOk = systemOk(sf);
  const snOk = systemOk(sn);
  const spOk = systemOk(sp);

  const sfStatus = sfOk ? "OK" : "error";
  const snStatus = snOk ? "OK" : "error";
  const spStatus = spOk ? "OK" : "error";

  // Salesforce insight
  const revenueLine = sfOk && sf?.atRiskSummary
    ? `At-risk opportunities: ${sf.atRiskSummary.opportunityCount} deal(s), total ~$${sf.atRiskSummary.totalAmount}.`
    : `We can’t see at-risk pipeline right now (visibility blind spot).`;

  // ServiceNow insight
  const itLine = snOk
    ? `High-priority open: ${sn.totalHighPriority}. (P1: ${sn.byPriority?.find(x => x.priority === "1")?.count ?? "?"}, P2: ${sn.byPriority?.find(x => x.priority === "2")?.count ?? "?"})`
    : `We can’t pull incident health right now (visibility blind spot).`;

  // SharePoint insight
  const spLine = spOk
    ? (sp.answer ? sp.answer : "No matching docs found for this phrasing.")
    : `We can’t pull knowledge signals right now (visibility blind spot).`;

  return `
**Leadership view (today):**
- One view across revenue risk, IT stability, and execution knowledge — with traceability by system.
- When any source fails, that itself becomes a leadership risk (visibility blind spot).

**Question asked:** ${question}

**Top 3 issues + business impact:**
1) **Revenue risk on strategic account** *(Source: Salesforce)*
- Impact: ${revenueLine}
- Next move (today): Ask the account owner for a save-plan + exec sponsor call today.

2) **High-priority incident load** *(Source: ServiceNow)*
- Impact: ${itLine}
- Next move (today): Demand a 24-hour stabilization plan: owners, top root causes, and ETA to restore service health.

3) **Execution / knowledge signal** *(Source: SharePoint)*
- Impact: ${spLine}
- Next move (today): Ask using an exact file name: **EBC_Account_Health_Risk.docx**, **IT_Operations_Weekly_Report.docx**, **Sales_Risk_Accounts_List.docx**.

**System status:** Salesforce: ${sfStatus} | ServiceNow: ${snStatus} | SharePoint: ${spStatus}
`.trim();
}

// -------------------- Handler --------------------

export default async function handler(req, res) {
  setCors(res);

  if (req.method === "OPTIONS") {
    res.statusCode = 200;
    return res.end();
  }

  if (req.method !== "POST") {
    return json(res, 405, { error: 'Use POST with JSON body {"question":"..."}' });
  }

  const question = safeString(req.body?.question);
  if (!question) {
    return json(res, 400, { error: 'Missing "question" in request body.' });
  }

  // Gather data in parallel; never throw the whole request.
  const [serviceNow, salesforce, sharePoint] = await Promise.all([
    fetchServiceNowSummary().catch((e) => ({ source: "ServiceNow", error: maskError(e) })),
    fetchSalesforceSummary().catch((e) => ({ source: "Salesforce", error: maskError(e) })),
    fetchSharePointSummary(question).catch((e) => ({ source: "SharePoint", error: maskError(e) })),
  ]);

  // Try Gemini; if fails, fallback formatting.
  const gemini = await callGeminiLeadershipSummary({
    question,
    serviceNow,
    salesforce,
    sharePoint,
  });

  const combinedAnswer = gemini.used && gemini.text
    ? gemini.text
    : formatFallbackAnswer(question, serviceNow, salesforce, sharePoint);

  return json(res, 200, {
    question,
    combinedAnswer,
    sources: {
      serviceNow,
      salesforce,
      sharePoint,
    },
    generatedAt: new Date().toISOString(),
    gemini: gemini.used
      ? { used: true, model: gemini.model }
      : { used: false, error: gemini.error || "Gemini not used" },
  });
}
