// /api/txi-dashboard.js
// Combined leadership view across ServiceNow, Salesforce, and SharePoint.

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
  // Self base URL for calling our own /api/chat-sp
  SELF_BASE_URL,
} = process.env;

export default async function handler(req, res) {
  // CORS + allowed methods (for browser + Postman)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    res.status(200).end();
    return;
  }

  // Accept GET ?question= and POST {question:""}
  let question;
  if (req.method === "POST") {
    question = req.body?.question;
  } else if (req.method === "GET") {
    question = req.query?.question;
  } else {
    res.setHeader("Allow", ["GET", "POST", "OPTIONS"]);
    res
      .status(405)
      .json({
        error: 'Use GET ?question=... or POST JSON body { "question": "..." }',
      });
    return;
  }

  if (!question || typeof question !== "string") {
    res
      .status(400)
      .json({ error: 'Missing "question" field in body or query string.' });
    return;
  }

  try {
    const [sn, sf, sp] = await Promise.all([
      getServiceNowSummary(),
      getSalesforceSummary(),
      getSharePointSummary(question),
    ]);

    const combinedAnswer = buildCombinedAnswer(question, sn, sf, sp);

    res.status(200).json({
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
    console.error("txi-dashboard error", err);
    res.status(500).json({ error: "Failed to generate combined answer." });
  }
}

/* -------------------- ServiceNow -------------------- */

async function getServiceNowSummary() {
  if (!SN_BASE_URL || !SN_USER || !SN_PASS) {
    return { source: "ServiceNow", error: "Missing ServiceNow env vars." };
  }

  const base = SN_BASE_URL.replace(/\/$/, "");
  const url = `${base}/api/dtp/schindler_txi/incident_summary`;

  try {
    const resp = await fetch(url, {
      method: "GET",
      headers: {
        Authorization:
          "Basic " + Buffer.from(`${SN_USER}:${SN_PASS}`).toString("base64"),
        Accept: "application/json",
      },
    });

    if (!resp.ok) {
      const text = await resp.text();
      return {
        source: "ServiceNow",
        error: `ServiceNow API error ${resp.status}: ${text}`,
      };
    }

    const data = await resp.json();
    // Example: { source:"ServiceNow", totalHighPriority, byPriority:[...], ebcIncidents:[...] }
    return { source: "ServiceNow", ...data };
  } catch (err) {
    return { source: "ServiceNow", error: String(err) };
  }
}

/* -------------------- Salesforce -------------------- */

async function getSalesforceSummary() {
  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN) {
    return { source: "Salesforce", error: "Missing Salesforce env vars." };
  }

  const conn = new jsforce.Connection({
    loginUrl: SF_LOGIN_URL || "https://login.salesforce.com",
  });

  try {
    await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);

    // 1. EBC account (checkbox Is_EBC_Account__c) – you have this now
    const ebcRes = await conn.query(
      "SELECT Id, Name, Industry, CustomerPriority__c, Active__c " +
        "FROM Account WHERE Is_EBC_Account__c = true LIMIT 1"
    );
    const ebc = ebcRes.records?.[0] || null;

    let atRiskOpps = [];

    if (ebc) {
      // 2. At-risk opportunities under that account
      const oppRes = await conn.query(
        `SELECT Id, Name, Amount, CloseDate, StageName, Risk_Flag__c, At_Risk__c ` +
          `FROM Opportunity ` +
          `WHERE AccountId = '${ebc.Id}' ` +
          `AND (Risk_Flag__c = true OR At_Risk__c = true)`
      );

      atRiskOpps = (oppRes.records || []).map((o) => ({
        id: o.Id,
        name: o.Name,
        amount: o.Amount,
        closeDate: o.CloseDate,
        stage: o.StageName,
      }));
    }

    return {
      source: "Salesforce",
      ebcAccount: ebc
        ? {
            id: ebc.Id,
            name: ebc.Name,
            industry: ebc.Industry,
            priority: ebc.CustomerPriority__c,
            active: ebc.Active__c,
          }
        : null,
      atRiskOpportunities: atRiskOpps,
    };
  } catch (err) {
    return { source: "Salesforce", error: String(err) };
  }
}

/* -------------------- SharePoint -------------------- */

async function getSharePointSummary(question) {
  // Call our own /api/chat-sp, which you’ve already wired to SharePoint+Gemini
  const base = (SELF_BASE_URL || "").trim() || "https://ai-bot-backend-black.vercel.app";

  try {
    const resp = await fetch(`${base.replace(/\/$/, "")}/api/chat-sp`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        question:
          question ||
          "Do we have any documents related to budget or deals?",
      }),
    });

    if (!resp.ok) {
      const text = await resp.text();
      return {
        source: "SharePoint",
        error: `/api/chat-sp error ${resp.status}: ${text}`,
      };
    }

    const data = await resp.json();
    return {
      source: "SharePoint",
      answer: data.answer,
      usedFiles: data.usedFiles || [],
    };
  } catch (err) {
    return { source: "SharePoint", error: String(err) };
  }
}

/* -------------------- Combined answer builder -------------------- */

function buildCombinedAnswer(question, sn, sf, sp) {
  const lines = [];

  lines.push(
    "Here is a leadership view of today’s biggest risks and customer impacts:"
  );
  lines.push("");

  // 1) Salesforce – money risk
  if (sf && !sf.error && sf.ebcAccount && sf.atRiskOpportunities?.length) {
    const total = sf.atRiskOpportunities.reduce(
      (sum, o) => sum + (Number(o.amount) || 0),
      0
    );
    lines.push(
      "1. At-risk opportunities with a high-priority EBC customer (Source: Salesforce)"
    );
    lines.push(
      `   - Impact: Potential loss of ~$${total.toLocaleString()} with ${sf.ebcAccount.name}, a high-priority ${sf.ebcAccount.industry} customer.`
    );
    lines.push("   - At-risk deals:");
    sf.atRiskOpportunities.forEach((o) => {
      lines.push(
        `     • ${o.name} – $${(o.amount ?? "").toLocaleString?.() ||
          o.amount} closing on ${o.closeDate} (stage: ${o.stage})`
      );
    });
    lines.push("");
  } else if (sf && sf.error) {
    lines.push(
      "1. Limited visibility into key customer accounts and sales risks (Source: Salesforce)"
    );
    lines.push(`   - Impact: ${sf.error}`);
    lines.push("");
  }

  // 2) ServiceNow – IT risk
  if (sn && !sn.error) {
    lines.push(
      "2. High-priority IT incident load and potential service risk (Source: ServiceNow)"
    );
    if (typeof sn.totalHighPriority === "number") {
      lines.push(
        `   - Impact: There are ${sn.totalHighPriority} high-priority incidents open.`
      );
    }
    if (Array.isArray(sn.byPriority)) {
      const summary = sn.byPriority
        .map((p) => `P${p.priority}: ${p.count}`)
        .join(", ");
      lines.push(`   - Distribution by priority: ${summary}.`);
    }
    lines.push(
      "   - Leadership risk: Without strong ownership, these incidents can directly hit uptime, employee productivity, and customer experience."
    );
    lines.push("");
  } else if (sn && sn.error) {
    lines.push(
      "2. Lack of IT operational visibility due to ServiceNow data issues (Source: ServiceNow)"
    );
    lines.push(`   - Impact: ${sn.error}`);
    lines.push("");
  }

  // 3) SharePoint – knowledge / execution risk
  if (sp && !sp.error && sp.answer) {
    lines.push(
      "3. Knowledge and execution risks from scattered documentation (Source: SharePoint)"
    );
    lines.push("   - The SharePoint assistant found relevant material:");
    lines.push(`     "${sp.answer.replace(/\s+/g, " ").trim()}"`);
    if (Array.isArray(sp.usedFiles) && sp.usedFiles.length) {
      const names = sp.usedFiles.map((f) => f.name).join(", ");
      lines.push(`   - Key files referenced: ${names}.`);
    }
    lines.push("");
  } else if (sp && sp.error) {
    lines.push(
      "3. Lack of collaboration and document visibility (Source: SharePoint)"
    );
    lines.push(`   - Impact: ${sp.error}`);
    lines.push("");
  }

  lines.push(
    "Overall, the biggest near-term risk is revenue leakage from at-risk EBC opportunities, amplified by IT visibility gaps and incomplete collaboration signals."
  );

  return lines.join("\n");
}
