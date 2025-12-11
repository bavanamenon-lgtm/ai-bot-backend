// /api/txi-dashboard.js
// Combined leadership view across ServiceNow, Salesforce, and SharePoint

import fetch from "node-fetch";
import jsforce from "jsforce";

const {
  // ServiceNow
  SN_BASE_URL,
  SN_USER,
  SN_PASS,

  // Salesforce
  SF_USERNAME,
  SF_PASSWORD,
  SF_TOKEN,
  SF_LOGIN_URL,

  // SharePoint assistant (optional)
  SP_CHAT_URL,
} = process.env;

/**
 * Utility: safe JSON fetch that never throws, it always returns an object
 */
async function safeJsonFetch(url, options = {}) {
  try {
    const res = await fetch(url, options);
    const text = await res.text();

    let data = null;
    try {
      data = text ? JSON.parse(text) : null;
    } catch (e) {
      // Non-JSON response
      data = null;
    }

    return {
      ok: res.ok,
      status: res.status,
      data,
      raw: text,
    };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: err.message || String(err),
    };
  }
}

/**
 * 1) ServiceNow – incident summary
 */
async function fetchServiceNowSummary() {
  if (!SN_BASE_URL || !SN_USER || !SN_PASS) {
    return {
      source: "ServiceNow",
      error: "ServiceNow environment variables are not configured.",
    };
  }

  const url = `${SN_BASE_URL.replace(/\/$/, "")}/api/dtp/schindler_txi/incident_summary`;

  const res = await safeJsonFetch(url, {
    method: "GET",
    headers: {
      Authorization:
        "Basic " +
        Buffer.from(`${SN_USER}:${SN_PASS}`).toString("base64"),
      "Content-Type": "application/json",
    },
  });

  if (!res.ok) {
    return {
      source: "ServiceNow",
      error:
        res.data?.error?.message ||
        `ServiceNow HTTP ${res.status}`,
      raw: res.raw,
    };
  }

  // Expected shape: { totalHighPriority, byPriority, ebcIncidents, ... }
  return {
    source: "ServiceNow",
    ...(res.data || {}),
  };
}

/**
 * 2) Salesforce – EBC account + at-risk opportunities
 */
async function fetchSalesforceSummary() {
  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN) {
    return {
      source: "Salesforce",
      error: "Salesforce environment variables are not configured.",
    };
  }

  const conn = new jsforce.Connection({
    loginUrl: SF_LOGIN_URL || "https://login.salesforce.com",
  });

  try {
    await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);
  } catch (err) {
    return {
      source: "Salesforce",
      error: `Salesforce login failed: ${err.message || err}`,
    };
  }

  try {
    // 2.1 EBC account – you can refine this query later if needed
    const accountQuery =
      "SELECT Id, Name, Industry, Rating FROM Account " +
      "WHERE Name = 'EBC HQ' LIMIT 1";
    const acctResult = await conn.query(accountQuery);
    const ebcAccount = acctResult.records?.[0] || null;

    // 2.2 At-risk opportunities for that account
    let atRiskOpportunities = [];
    let atRiskSummary = null;

    if (ebcAccount) {
      const oppQuery =
        "SELECT Id, Name, Amount, StageName, CloseDate, Probability, " +
        "Risk_Flag__c, At_Risk__c " +
        "FROM Opportunity " +
        `WHERE AccountId = '${ebcAccount.Id}' ` +
        "AND Risk_Flag__c = true " +
        "AND At_Risk__c = true";

      const oppResult = await conn.query(oppQuery);
      atRiskOpportunities = (oppResult.records || []).map((o) => ({
        id: o.Id,
        name: o.Name,
        amount: o.Amount,
        stage: o.StageName,
        closeDate: o.CloseDate,
        probability: o.Probability,
      }));

      const totalAmount = atRiskOpportunities.reduce(
        (sum, o) => sum + (o.amount || 0),
        0
      );

      atRiskSummary = {
        opportunityCount: atRiskOpportunities.length,
        totalAmount,
      };
    }

    return {
      source: "Salesforce",
      ebcAccount,
      atRiskSummary,
      atRiskOpportunities,
    };
  } catch (err) {
    return {
      source: "Salesforce",
      error: err.message || String(err),
    };
  }
}

/**
 * 3) SharePoint – call your existing /api/chat-sp endpoint (optional)
 * If SP_CHAT_URL is not configured, we just report that as a soft error.
 */
async function fetchSharePointSummary(question) {
  if (!SP_CHAT_URL) {
    return {
      source: "SharePoint",
      error: "SP_CHAT_URL env var not configured.",
    };
  }

  const res = await safeJsonFetch(SP_CHAT_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ question }),
  });

  if (!res.ok) {
    return {
      source: "SharePoint",
      error:
        res.data?.error ||
        `SharePoint assistant HTTP ${res.status}`,
      raw: res.raw,
    };
  }

  // Assume your /api/chat-sp returns { answer, usedFiles, candidateFiles, ... }
  return {
    source: "SharePoint",
    ...(res.data || {}),
  };
}

/**
 * Build a plain-English leadership answer WITHOUT calling Gemini,
 * so we don’t add another failure point. You can plug Gemini back in later
 * if you really want the flowery version.
 */
function buildLeadershipAnswer(question, { serviceNow, salesforce, sharePoint }) {
  const lines = [];

  lines.push("Here is a leadership view of today’s biggest risks and impacts:\n");

  // 1) Salesforce
  if (salesforce?.error) {
    lines.push(
      `1. Salesforce visibility issue\n   - Impact: ${salesforce.error}\n`
    );
  } else if (salesforce?.ebcAccount && salesforce?.atRiskSummary) {
    const acc = salesforce.ebcAccount;
    const sum = salesforce.atRiskSummary;

    lines.push(
      `1. Revenue risk on key customer account (Source: Salesforce)\n` +
        `   - Account: ${acc.Name} (Industry: ${acc.Industry || "n/a"}, Rating: ${
          acc.Rating || "n/a"
        })\n` +
        `   - Open at-risk opportunities: ${sum.opportunityCount} deal(s) with total value around $${sum.totalAmount}\n` +
        `   - Leadership impact: Losing or delaying these deals directly impacts revenue and the relationship with this strategic account.\n`
    );
  } else {
    lines.push(
      `1. Limited insight into key accounts and at-risk revenue (Source: Salesforce)\n` +
        `   - Impact: No EBC account or at-risk opportunities could be identified from the current data snapshot.\n`
    );
  }

  // 2) ServiceNow
  if (serviceNow?.error) {
    lines.push(
      `2. IT operations visibility gap (Source: ServiceNow)\n` +
        `   - Impact: ${serviceNow.error}. This prevents reliable tracking of incidents and service health.\n`
    );
  } else if (typeof serviceNow?.totalHighPriority === "number") {
    const total = serviceNow.totalHighPriority;
    const byPri = serviceNow.byPriority || [];
    const breakdown = byPri
      .map((p) => `P${p.priority}: ${p.count}`)
      .join(", ");

    lines.push(
      `2. High-priority IT incident load (Source: ServiceNow)\n` +
        `   - High-priority incidents open: ${total}\n` +
        `   - Distribution by priority: ${breakdown || "n/a"}\n` +
        `   - Leadership impact: Persistent high-priority incidents put uptime, employee productivity and customer experience at risk.\n`
    );
  } else {
    lines.push(
      `2. Unknown IT incident status (Source: ServiceNow)\n` +
        `   - Impact: Incident summary data was not available, so we cannot quantify current IT risk.\n`
    );
  }

  // 3) SharePoint
  if (sharePoint?.error) {
    lines.push(
      `3. Collaboration / knowledge visibility gap (Source: SharePoint)\n` +
        `   - Impact: ${sharePoint.error}\n`
    );
  } else if (sharePoint?.answer) {
    lines.push(
      `3. Knowledge insights from SharePoint (Source: SharePoint)\n` +
        `   - Assistant summary: ${sharePoint.answer}\n`
    );
  } else {
    lines.push(
      `3. No specific SharePoint insight surfaced for this question.\n`
    );
  }

  return lines.join("\n");
}

/**
 * API handler
 */
export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res
      .status(405)
      .json({ error: 'Use POST with JSON body { "question": "..." }' });
  }

  const body = req.body || {};
  const question = body.question;

  if (!question || typeof question !== "string") {
    return res
      .status(400)
      .json({ error: 'Missing "question" in request body or not a string.' });
  }

  try {
    const [serviceNow, salesforce, sharePoint] = await Promise.all([
      fetchServiceNowSummary(),
      fetchSalesforceSummary(),
      fetchSharePointSummary(question),
    ]);

    const combinedAnswer = buildLeadershipAnswer(question, {
      serviceNow,
      salesforce,
      sharePoint,
    });

    return res.status(200).json({
      question,
      combinedAnswer,
      sources: {
        serviceNow,
        salesforce,
        sharePoint,
      },
      generatedAt: new Date().toISOString(),
    });
  } catch (err) {
    // THIS is the “no more 500s” catch-all:
    console.error("Fatal error in /api/txi-dashboard:", err);

    return res.status(200).json({
      question,
      combinedAnswer:
        "Could not generate a full answer due to an internal error in the TXI dashboard backend.",
      error: err.message || String(err),
      sources: {
        serviceNow: null,
        salesforce: null,
        sharePoint: null,
      },
      generatedAt: new Date().toISOString(),
    });
  }
}
