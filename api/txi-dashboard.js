// api/txi-dashboard.js
// Schindler TXI Dashboard: ServiceNow + Salesforce + SharePoint -> (Optional) Gemini polish
// Always returns a result (never fail whole response because one system fails).

import fetch from "node-fetch";
import jsforce from "jsforce";

const {
  // ServiceNow
  SN_TXI_URL,
  SN_USERNAME,
  SN_PASSWORD,

  // Salesforce
  SF_USERNAME,
  SF_PASSWORD,
  SF_TOKEN,
  SF_LOGIN_URL,

  // SharePoint (calls our own chat-sp endpoint)
  SP_CHAT_URL,

  // Gemini
  GEMINI_API_KEY,
  GEMINI_MODEL,
} = process.env;

function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
}

function safeNum(n) {
  const x = Number(n);
  return Number.isFinite(x) ? x : 0;
}

function shortMoneyUSD(amount) {
  const a = safeNum(amount);
  return a.toLocaleString("en-US");
}

function buildSharePointQuery(question = "") {
  const q = (question || "").toLowerCase();

  // map leadership phrases -> file keywords
  if (q.includes("deal") || q.includes("revenue") || q.includes("pipeline") || q.includes("budget") || q.includes("opportun")) {
    return `Search SharePoint for "Deals Pipeline" OR "Execution Plan" OR "Action Items" and summarise the most relevant file.`;
  }
  if (q.includes("incident") || q.includes("outage") || q.includes("service") || q.includes("sla")) {
    return `Search SharePoint for "incident" OR "postmortem" OR "SLA" OR "major incident" and summarise the most relevant file.`;
  }
  return `Search SharePoint for "Execution Plan" OR "Action Items" and summarise the most relevant file.`;
}

async function fetchServiceNow() {
  if (!SN_TXI_URL || !SN_USERNAME || !SN_PASSWORD) {
    return { source: "ServiceNow", error: "ServiceNow env vars not configured (SN_TXI_URL / SN_USERNAME / SN_PASSWORD)." };
  }

  const auth = Buffer.from(`${SN_USERNAME}:${SN_PASSWORD}`).toString("base64");

  const resp = await fetch(SN_TXI_URL, {
    method: "GET",
    headers: {
      "Authorization": `Basic ${auth}`,
      "Accept": "application/json",
    },
  });

  const text = await resp.text();
  let data = null;
  try { data = JSON.parse(text); } catch (e) {}

  if (!resp.ok) {
    return {
      source: "ServiceNow",
      error: `ServiceNow API error (${resp.status})`,
      raw: (data || text || "").toString().slice(0, 800),
    };
  }

  // Expecting your endpoint already returns: totalHighPriority, byPriority, etc.
  return data && data.serviceNow ? data.serviceNow : data;
}

async function fetchSalesforce() {
  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN || !SF_LOGIN_URL) {
    return { source: "Salesforce", error: "Salesforce env vars not configured (SF_USERNAME / SF_PASSWORD / SF_TOKEN / SF_LOGIN_URL)." };
  }

  const conn = new jsforce.Connection({ loginUrl: SF_LOGIN_URL });
  await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);

  // You said you already created these fields:
  // Account: Is_EBC_Account__c
  // Opportunity: Risk_Flag__c (checkbox), Amount, StageName, CloseDate, Probability
  const ebcAccounts = await conn.query(`
    SELECT Id, Name, Industry, Rating
    FROM Account
    WHERE Is_EBC_Account__c = true
    LIMIT 1
  `);

  const ebcAccount = (ebcAccounts.records && ebcAccounts.records[0]) ? ebcAccounts.records[0] : null;

  let atRiskOpps = [];
  if (ebcAccount) {
    const opps = await conn.query(`
      SELECT Id, Name, Amount, StageName, CloseDate, Probability
      FROM Opportunity
      WHERE AccountId = '${ebcAccount.Id}'
        AND Risk_Flag__c = true
      ORDER BY CloseDate ASC
      LIMIT 5
    `);
    atRiskOpps = opps.records || [];
  }

  const totalAmount = atRiskOpps.reduce((s, o) => s + safeNum(o.Amount), 0);

  return {
    source: "Salesforce",
    ebcAccount: ebcAccount
      ? { id: ebcAccount.Id, name: ebcAccount.Name, industry: ebcAccount.Industry, rating: ebcAccount.Rating }
      : null,
    atRiskSummary: { opportunityCount: atRiskOpps.length, totalAmount },
    atRiskOpportunities: atRiskOpps.map(o => ({
      id: o.Id,
      name: o.Name,
      amount: safeNum(o.Amount),
      stage: o.StageName,
      closeDate: o.CloseDate,
      probability: safeNum(o.Probability),
    })),
  };
}

async function fetchSharePoint(question) {
  if (!SP_CHAT_URL) {
    return { source: "SharePoint", error: "SP_CHAT_URL env var not configured." };
  }

  const spQuestion = buildSharePointQuery(question);

  const resp = await fetch(SP_CHAT_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ question: spQuestion }),
  });

  const text = await resp.text();
  let data = null;
  try { data = JSON.parse(text); } catch (e) {}

  if (!resp.ok || !data) {
    return {
      source: "SharePoint",
      error: `SharePoint /api/chat-sp error (${resp.status})`,
      raw: (text || "").slice(0, 800),
    };
  }

  // chat-sp returns { answer, usedFiles, candidateFiles } in your latest working version
  return {
    source: "SharePoint",
    answer: data.answer || "No answer from SharePoint.",
    usedFiles: data.usedFiles || [],
    candidateFiles: data.candidateFiles || [],
  };
}

function buildDeterministicExecutiveAnswer({ question, sn, sf, sp }) {
  const bullets = [];

  // 1) Salesforce
  if (sf?.error) {
    bullets.push({
      title: "Revenue / customer risk visibility gap",
      source: "Salesforce",
      impact: `We couldn’t pull strategic account + at-risk pipeline due to an error: ${sf.error}`,
      action: "Fix the Salesforce query/permissions and re-run. Until then, leadership is flying blind on revenue risk.",
    });
  } else if (sf?.ebcAccount) {
    const amt = shortMoneyUSD(sf.atRiskSummary?.totalAmount || 0);
    bullets.push({
      title: "Revenue risk on a strategic account",
      source: "Salesforce",
      impact: `${sf.ebcAccount.name} (Rating: ${sf.ebcAccount.rating || "N/A"}) has ${sf.atRiskSummary.opportunityCount} at-risk opportunity(ies) worth ~$${amt}.`,
      action: "Ask the account owner for a save-plan today: next steps, blockers, and exec coverage for each at-risk deal.",
    });
  } else {
    bullets.push({
      title: "No strategic (EBC) account found in sample data",
      source: "Salesforce",
      impact: "The dashboard didn’t find any EBC account in Salesforce, so it can’t highlight strategic customer risk.",
      action: "Create or tag one EBC account and link 1–2 at-risk opportunities to demonstrate the value properly.",
    });
  }

  // 2) ServiceNow
  if (sn?.error) {
    bullets.push({
      title: "IT stability visibility gap",
      source: "ServiceNow",
      impact: `We couldn’t pull incident health due to: ${sn.error}`,
      action: "Fix ServiceNow auth/env vars and confirm the API works from Vercel (GET).",
    });
  } else {
    const p1 = (sn.byPriority || []).find(x => String(x.priority) === "1")?.count ?? null;
    const p2 = (sn.byPriority || []).find(x => String(x.priority) === "2")?.count ?? null;
    bullets.push({
      title: "High priority incident load",
      source: "ServiceNow",
      impact: `High-priority incidents open: ${sn.totalHighPriority ?? "N/A"} (P1: ${p1 ?? "?"}, P2: ${p2 ?? "?"}).`,
      action: "Demand a 24-hour stabilization plan: top 5 root causes, owners, and expected time to restore service health.",
    });
  }

  // 3) SharePoint
  if (sp?.error) {
    bullets.push({
      title: "Knowledge / execution visibility gap",
      source: "SharePoint",
      impact: `We couldn’t pull relevant docs due to: ${sp.error}`,
      action: "Fix SP_CHAT_URL + confirm Graph permissions + verify the correct site/library scope.",
    });
  } else {
    const used = (sp.usedFiles && sp.usedFiles.length) ? sp.usedFiles.map(f => f.name).join(", ") : "No file matched";
    bullets.push({
      title: "Execution context from internal documents",
      source: "SharePoint",
      impact: `Docs used: ${used}.`,
      action: "Use this to validate whether sales/operations actions already exist, and whether teams are following the plan.",
    });
  }

  // Headline (single sentence)
  const headline = "One view across revenue risk, IT stability, and execution knowledge — with traceability by system.";

  // Build final
  let out = `**Leadership view (TXI POC):** ${headline}\n\n`;
  out += `**Question asked:** ${question}\n\n`;
  out += `**Top 3 issues + what to do next:**\n`;
  bullets.slice(0, 3).forEach((b, i) => {
    out += `\n${i + 1}) **${b.title}** *(Source: ${b.source})*\n`;
    out += `- **Impact:** ${b.impact}\n`;
    out += `- **Next action:** ${b.action}\n`;
  });

  out += `\n\n**System status:** ServiceNow: ${sn?.error ? "error" : "OK"} | Salesforce: ${sf?.error ? "error" : "OK"} | SharePoint: ${sp?.error ? "error" : "OK"}\n`;
  out += `**Note:** This is a POC using sample data + limited scope; production would add stricter masking, access controls, and governance.\n`;

  return out;
}

async function callGeminiPolish({ question, deterministicAnswer, rawSources }) {
  if (!GEMINI_API_KEY) return { ok: false, error: "GEMINI_API_KEY not configured" };

  const model = GEMINI_MODEL || "gemini-2.5-flash";
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GEMINI_API_KEY}`;

  const prompt = `
You are writing an executive-ready answer for a leadership dashboard.

User question:
${question}

Raw signals (JSON):
${JSON.stringify(rawSources).slice(0, 12000)}

Draft answer (deterministic):
${deterministicAnswer}

Rewrite the draft answer to be:
- extremely clear for leadership
- short, structured, skimmable
- no hallucination: ONLY use numbers present in raw JSON
- keep "Top 3 issues" format
- keep system status line
- do NOT mention internal implementation terms (Vercel, Postman, env vars) unless it's explicitly an error cause shown in rawSources
`.trim();

  const body = { contents: [{ parts: [{ text: prompt }] }] };

  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  const raw = await resp.text();
  let data = null;
  try { data = JSON.parse(raw); } catch (e) {}

  if (!resp.ok || !data) {
    return { ok: false, error: `Gemini error (${resp.status})`, raw: raw.slice(0, 600) };
  }

  const textOut =
    data?.candidates?.[0]?.content?.parts?.[0]?.text ||
    null;

  return textOut ? { ok: true, text: textOut } : { ok: false, error: "Gemini returned empty text" };
}

export default async function handler(req, res) {
  setCors(res);

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: 'Use POST with JSON body { "question": "..." }' });

  try {
    const { question } = req.body || {};
    if (!question || typeof question !== "string") {
      return res.status(400).json({ error: 'Missing "question" in request body.' });
    }

    // pull all three in parallel
    const [sn, sf, sp] = await Promise.allSettled([
      fetchServiceNow(),
      fetchSalesforce(),
      fetchSharePoint(question),
    ]);

    const snVal = sn.status === "fulfilled" ? sn.value : { source: "ServiceNow", error: String(sn.reason) };
    const sfVal = sf.status === "fulfilled" ? sf.value : { source: "Salesforce", error: String(sf.reason) };
    const spVal = sp.status === "fulfilled" ? sp.value : { source: "SharePoint", error: String(sp.reason) };

    const deterministic = buildDeterministicExecutiveAnswer({
      question,
      sn: snVal,
      sf: sfVal,
      sp: spVal,
    });

    // Gemini is optional "polish"
    const rawSources = { serviceNow: snVal, salesforce: sfVal, sharePoint: spVal };
    const polished = await callGeminiPolish({
      question,
      deterministicAnswer: deterministic,
      rawSources,
    });

    const finalAnswer = polished.ok ? polished.text : deterministic;

    return res.status(200).json({
      question,
      combinedAnswer: finalAnswer,
      sources: rawSources,
      generatedAt: new Date().toISOString(),
      gemini: polished.ok ? { used: true, model: GEMINI_MODEL || "gemini-2.5-flash" } : { used: false, error: polished.error },
    });
  } catch (err) {
    return res.status(500).json({
      error: "Internal error in /api/txi-dashboard",
      details: String(err),
    });
  }
}
