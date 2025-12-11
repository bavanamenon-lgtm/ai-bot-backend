// api/txi-dashboard.js
// Combined leadership view across ServiceNow, Salesforce and SharePoint.

import fetch from "node-fetch";
import jsforce from "jsforce";

const {
  SN_BASE_URL,
  SN_USER,
  SN_PASS,

  SF_USERNAME,
  SF_PASSWORD,
  SF_TOKEN,
  SF_LOGIN_URL,

  // Full URL of your SharePoint Q&A endpoint
  SP_CHAT_URL,

  GEMINI_API_KEY, // not used right now, but kept for later
} = process.env;

// -----------------------------
// Helper: ServiceNow summary
// -----------------------------
async function getServiceNowSummary() {
  const url = `${SN_BASE_URL}/api/dtp/schindler_txi/incident_summary`;

  const auth = Buffer.from(`${SN_USER}:${SN_PASS}`).toString("base64");

  const resp = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Basic ${auth}`,
      Accept: "application/json",
    },
  });

  if (!resp.ok) {
    throw new Error(`ServiceNow API error: ${resp.status} ${resp.statusText}`);
  }

  const data = await resp.json();
  return {
    source: "ServiceNow",
    ...data,
  };
}

// -----------------------------
// Helper: Salesforce summary
// -----------------------------
async function getSalesforceSummary() {
  const conn = new jsforce.Connection({
    loginUrl: SF_LOGIN_URL || "https://login.salesforce.com",
  });

  await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);

  // 1) Find the EBC HQ account by name (no custom fields!)
  const accountResult = await conn.query(
    "SELECT Id, Name, Industry, Rating " +
      "FROM Account " +
      "WHERE Name = 'EBC HQ' " +
      "LIMIT 1"
  );

  if (!accountResult.records.length) {
    return {
      source: "Salesforce",
      error: "No account called 'EBC HQ' found in Salesforce.",
    };
  }

  const account = accountResult.records[0];

  // 2) Find open / future opportunities on that account
  const oppResult = await conn.query(
    "SELECT Id, Name, Amount, StageName, CloseDate, Probability " +
      "FROM Opportunity " +
      `WHERE AccountId = '${account.Id}' ` +
      "AND IsClosed = false " +
      "AND CloseDate >= TODAY"
  );

  const opps = oppResult.records || [];

  // Treat all open opps as “at risk” for the POC
  let totalAtRiskAmount = 0;
  opps.forEach((o) => {
    if (o.Amount) totalAtRiskAmount += o.Amount;
  });

  return {
    source: "Salesforce",
    ebcAccount: {
      id: account.Id,
      name: account.Name,
      industry: account.Industry,
      rating: account.Rating,
    },
    atRiskSummary: {
      opportunityCount: opps.length,
      totalAmount: totalAtRiskAmount,
    },
    atRiskOpportunities: opps.map((o) => ({
      id: o.Id,
      name: o.Name,
      amount: o.Amount,
      stage: o.StageName,
      closeDate: o.CloseDate,
      probability: o.Probability,
    })),
  };
}

// -----------------------------
// Helper: SharePoint / document insight
// -----------------------------
async function getSharePointSummary(question) {
  if (!SP_CHAT_URL) {
    return {
      source: "SharePoint",
      error: "SP_CHAT_URL env var not configured.",
    };
  }

  const resp = await fetch(SP_CHAT_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      // You can tune this sub-question later
      question:
        "Do we have any documents or plans related to risks, budget, or deals that leadership should know about?",
    }),
  });

  if (!resp.ok) {
    throw new Error(
      `SharePoint chat API error: ${resp.status} ${resp.statusText}`
    );
  }

  const data = await resp.json();
  return {
    source: "SharePoint",
    ...data,
  };
}

// -----------------------------
// Build a simple combined answer
// -----------------------------
function buildCombinedAnswer({ question, sn, sf, sp }) {
  let lines = [];

  lines.push("Here is a leadership view of today’s biggest risks and impacts:\n");

  // 1) Salesforce – revenue risk
  if (sf && !sf.error && sf.atRiskSummary) {
    const count = sf.atRiskSummary.opportunityCount;
    const amt = sf.atRiskSummary.totalAmount || 0;

    lines.push(
      `1. Revenue risk on key customer account (Source: Salesforce)\n` +
        `   - Account: ${sf.ebcAccount?.name || "N/A"} (Industry: ${
          sf.ebcAccount?.industry || "N/A"
        }, Rating: ${sf.ebcAccount?.rating || "N/A"})\n` +
        `   - Open opportunities: ${count} deal(s) with total value around $${amt.toLocaleString()}\n` +
        `   - Leadership impact: Losing or delaying these deals directly impacts revenue and the relationship with this strategic account.\n`
    );
  } else {
    lines.push(
      `1. Salesforce configuration gap (Source: Salesforce)\n` +
        `   - Impact: ${sf?.error || "Salesforce data is not yet available for this console."}\n`
    );
  }

  // 2) ServiceNow – high-priority incidents
  if (sn && !sn.error && typeof sn.totalHighPriority === "number") {
    lines.push(
      `2. High-priority IT incident load (Source: ServiceNow)\n` +
        `   - High-priority incidents open: ${sn.totalHighPriority}\n` +
        `   - Distribution by priority: ${sn.byPriority
          .map((p) => `P${p.priority}: ${p.count}`)
          .join(", ")}\n` +
        `   - Leadership impact: Persistent high-priority incidents put uptime, employee productivity and customer experience at risk.\n`
    );
  } else {
    lines.push(
      `2. IT visibility gap (Source: ServiceNow)\n` +
        `   - Impact: ${sn?.error || "Incident summary is not available; IT health is opaque at the moment."}\n`
    );
  }

  // 3) SharePoint – documentation / execution risk
  if (sp && !sp.error && sp.answer) {
    lines.push(
      `3. Knowledge and execution signals (Source: SharePoint)\n` +
        `   - Summary of most relevant document(s): ${sp.answer}\n` +
        `   - Leadership impact: These docs show how risks, budgets or deals are being tracked today; gaps here translate to execution risk.\n`
    );
  } else {
    lines.push(
      `3. Collaboration / knowledge visibility gap (Source: SharePoint)\n` +
        `   - Impact: ${sp?.error || "No relevant documents could be surfaced for this question."}\n`
    );
  }

  return lines.join("\n");
}

// -----------------------------
// API handler
// -----------------------------
export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res
      .status(405)
      .json({ error: 'Use POST with JSON body { "question": "..." }' });
  }

  try {
    const { question } = req.body || {};
    if (!question || typeof question !== "string") {
      return res.status(400).json({
        error: 'Missing "question" in request body or not a string.',
      });
    }

    // Run all three in parallel
    const [sn, sf, sp] = await Promise.all([
      getServiceNowSummary().catch((e) => ({ source: "ServiceNow", error: e.message })),
      getSalesforceSummary().catch((e) => ({ source: "Salesforce", error: e.message })),
      getSharePointSummary(question).catch((e) => ({
        source: "SharePoint",
        error: e.message,
      })),
    ]);

    const combinedAnswer = buildCombinedAnswer({ question, sn, sf, sp });

    return res.status(200).json({
      question,
      combinedAnswer,
      sources: {
        serviceNow: sn,
        salesforce: sf,
        sharePoint: sp,
      },
      generatedAt: new Date().toISOString(),
    });
  } catch (err) {
    console.error("txi-dashboard error:", err);
    return res.status(500).json({
      error: "Internal error in txi-dashboard.",
      detail: err.message,
    });
  }
}
