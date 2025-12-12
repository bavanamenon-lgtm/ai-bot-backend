// api/txi-dashboard.js
//
// Schindler TXI – Leadership Console backend
// - Fetches summary from ServiceNow TXI endpoint
// - Fetches EBC / at-risk data from Salesforce
// - Fetches doc insight from SharePoint AI assistant
// - Builds a clean leadership summary WITHOUT calling Gemini
//
// Assumes Node ESM + Vercel API route

import fetch from "node-fetch";
import jsforce from "jsforce";

const {
  // Salesforce
  SF_USERNAME,
  SF_PASSWORD,
  SF_TOKEN,
  SF_LOGIN_URL,

  // ServiceNow (TXI incident summary endpoint)
  SN_TXI_URL,
  SN_USERNAME,
  SN_PASSWORD,

  // SharePoint assistant (your existing /api/chat-sp endpoint URL)
  SP_CHAT_URL,
} = process.env;

// ---------- Helpers ----------

function basicAuthHeader(user, pass) {
  const token = Buffer.from(`${user}:${pass}`).toString("base64");
  return `Basic ${token}`;
}

// ---------- ServiceNow: TXI incident summary ----------

async function fetchServiceNowSummary() {
  if (!SN_TXI_URL || !SN_USERNAME || !SN_PASSWORD) {
    return {
      source: "ServiceNow",
      error:
        "ServiceNow env vars not configured (SN_TXI_URL / SN_USERNAME / SN_PASSWORD).",
    };
  }

  try {
    const resp = await fetch(SN_TXI_URL, {
      method: "GET",
      headers: {
        Authorization: basicAuthHeader(SN_USERNAME, SN_PASSWORD),
        "Content-Type": "application/json",
      },
    });

    const raw = await resp.text();
    let data;
    try {
      data = JSON.parse(raw);
    } catch (e) {
      return {
        source: "ServiceNow",
        error: "ServiceNow TXI endpoint returned non-JSON response.",
      };
    }

    if (!resp.ok) {
      return {
        source: "ServiceNow",
        error:
          data.error ||
          `ServiceNow API error (status ${resp.status || "unknown"})`,
      };
    }

    // Expecting shape:
    // {
    //   source: "ServiceNow",
    //   generatedAt: "...",
    //   totalHighPriority: 78,
    //   byPriority: [{priority:"1",count:72},...],
    //   ebcIncidents: [...]
    // }
    return data;
  } catch (err) {
    return {
      source: "ServiceNow",
      error: `ServiceNow fetch error: ${String(err)}`,
    };
  }
}

// ---------- Salesforce: EBC + at-risk opps ----------

async function fetchSalesforceSummary() {
  if (!SF_USERNAME || !SF_PASSWORD || ! SF_TOKEN || !SF_LOGIN_URL) {
    return {
      source: "Salesforce",
      error:
        "Salesforce env vars not configured (SF_USERNAME / SF_PASSWORD / SF_TOKEN / SF_LOGIN_URL).",
    };
  }

  const conn = new jsforce.Connection({
    loginUrl: SF_LOGIN_URL,
  });

  try {
    await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);

    // 1) Find the EBC HQ account (you already created this in sample data)
    const accResult = await conn.query(
      "SELECT Id, Name, Industry, Rating " +
        "FROM Account " +
        "WHERE Name = 'EBC HQ' " +
        "LIMIT 1"
    );

    const ebcAccount = accResult.records && accResult.records[0];
    if (!ebcAccount) {
      return {
        source: "Salesforce",
        ebcAccount: null,
        atRiskSummary: null,
        atRiskOpportunities: [],
        note: "No EBC HQ account found in Salesforce sample data.",
      };
    }

    // 2) At-risk opportunities for that account
    // You created custom checkbox "Risk Flag" on Opportunity.
    // API name assumed: Risk_Flag__c
    let oppQuery =
      "SELECT Id, Name, Amount, StageName, CloseDate, Probability " +
      "FROM Opportunity " +
      `WHERE AccountId = '${ebcAccount.Id}' ` +
      "AND Risk_Flag__c = true";

    const oppResult = await conn.query(oppQuery);

    const atRiskOpportunities = (oppResult.records || []).map((r) => ({
      id: r.Id,
      name: r.Name,
      amount: r.Amount,
      stage: r.StageName,
      closeDate: r.CloseDate,
      probability: r.Probability,
    }));

    const totalAmount = atRiskOpportunities.reduce(
      (sum, o) => sum + (o.amount || 0),
      0
    );

    return {
      source: "Salesforce",
      ebcAccount: {
        id: ebcAccount.Id,
        name: ebcAccount.Name,
        industry: ebcAccount.Industry,
        rating: ebcAccount.Rating,
      },
      atRiskSummary: {
        opportunityCount: atRiskOpportunities.length,
        totalAmount,
      },
      atRiskOpportunities,
    };
  } catch (err) {
    return {
      source: "Salesforce",
      error: `Salesforce API error: ${String(err)}`,
    };
  }
}

// ---------- SharePoint: ask existing /api/chat-sp ----------

async function fetchSharePointSummary(question) {
  if (!SP_CHAT_URL) {
    return {
      source: "SharePoint",
      error: "SP_CHAT_URL env var not configured.",
    };
  }

  try {
    const resp = await fetch(SP_CHAT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ question }),
    });

    const raw = await resp.text();
    let data;
    try {
      data = JSON.parse(raw);
    } catch (e) {
      return {
        source: "SharePoint",
        error: "SharePoint assistant returned non-JSON response.",
      };
    }

    if (!resp.ok) {
      return {
        source: "SharePoint",
        error:
          data.error ||
          `SharePoint assistant error (status ${resp.status || "unknown"})`,
      };
    }

    // Expected shape from /api/chat-sp:
    // {
    //   answer: "...",
    //   chosenFile: {...} (optional),
    //   sharePointResults: [...]
    // }
    return {
      source: "SharePoint",
      answer: data.answer || "",
      chosenFile: data.chosenFile || null,
      usedFiles: data.sharePointResults || [],
    };
  } catch (err) {
    return {
      source: "SharePoint",
      error: `SharePoint fetch error: ${String(err)}`,
    };
  }
}

// ---------- Build leadership summary (NO Gemini) ----------

function buildLeadershipSummary(question, sn, sf, sp) {
  const lines = [];

  lines.push(
    "Here is a leadership view of today’s biggest issues and business impact based on the available data:\n"
  );

  // 1) Salesforce – revenue / customer risk
  if (!sf || sf.error) {
    lines.push(
      `* **Revenue risk and customer visibility gap (Salesforce)**\n` +
        `  * Business impact: Salesforce data could not be reliably retrieved (${sf?.error || "unknown error"}). ` +
        `Leadership cannot clearly see at-risk revenue or key account health.\n`
    );
  } else if (
    sf.atRiskSummary &&
    sf.atRiskSummary.opportunityCount > 0 &&
    sf.ebcAccount
  ) {
    const total = sf.atRiskSummary.totalAmount || 0;
    lines.push(
      `* **Revenue risk on key customer account (Salesforce)**\n` +
        `  * Account: ${sf.ebcAccount.name} (Industry: ${sf.ebcAccount.industry || "N/A"}, Rating: ${
          sf.ebcAccount.rating || "N/A"
        })\n` +
        `  * Open at-risk opportunities: ${sf.atRiskSummary.opportunityCount} deal(s) with total value around $${total}\n` +
        `  * Leadership impact: Losing or delaying these deals directly impacts revenue and the relationship with this strategic account.\n`
    );
  } else {
    lines.push(
      `* **Limited insight into key customer accounts (Salesforce)**\n` +
        `  * Business impact: No at-risk opportunities were identified in the current sample data. ` +
        `This may mean low immediate risk, or it may indicate that risk flags and account health are not being tracked consistently.\n`
    );
  }

  // 2) ServiceNow – IT stability
  if (!sn || sn.error) {
    lines.push(
      `* **IT service visibility gap (ServiceNow)**\n` +
        `  * Business impact: ServiceNow data could not be retrieved (${sn?.error || "unknown error"}). ` +
        `Leadership lacks visibility into high-priority incidents, which can hide outages and user-impacting issues.\n`
    );
  } else {
    const totalHigh = sn.totalHighPriority ?? null;
    const byPriority = Array.isArray(sn.byPriority) ? sn.byPriority : [];

    let distText = "";
    if (byPriority.length) {
      const parts = byPriority.map(
        (p) => `P${p.priority}: ${p.count}`
      );
      distText = parts.join(", ");
    }

    lines.push(
      `* **High-priority IT incident load (ServiceNow)**\n` +
        `  * High-priority incidents open: ${
          totalHigh != null ? totalHigh : "N/A"
        }\n` +
        (distText
          ? `  * Distribution by priority: ${distText}\n`
          : "") +
        `  * Leadership impact: Persistent high-priority incidents put uptime, employee productivity and customer experience at risk.\n`
    );
  }

  // 3) SharePoint – knowledge / collaboration
  if (!sp || sp.error) {
    lines.push(
      `* **Collaboration / knowledge visibility gap (SharePoint)**\n` +
        `  * Impact: ${sp?.error || "SharePoint AI assistant is not reachable."} ` +
        `This prevents leadership from quickly surfacing key documents and execution signals tied to strategic initiatives.\n`
    );
  } else if (sp.answer) {
    lines.push(
      `* **Knowledge insights from SharePoint (documents)**\n` +
        `  * Assistant summary: ${sp.answer.trim()}\n`
    );
  } else {
    lines.push(
      `* **Limited document insights from SharePoint**\n` +
        `  * Impact: No strong document signals were surfaced in this query. ` +
        `Either relevant documents are not consistently tagged/named, or the current library does not yet contain rich execution data.\n`
    );
  }

  lines.push(
    "\n---\n" +
      "This is a Proof of Concept using limited sample data, but it shows how a single leadership question can pull together revenue risk (Salesforce), IT stability (ServiceNow), and knowledge/collaboration signals (SharePoint) into one executive view."
  );

  return lines.join("");
}

// ---------- Main handler ----------

export default async function handler(req, res) {
  if (req.method === "OPTIONS") {
    // Allow preflight if ever used cross-origin
    res.status(200).end();
    return;
  }

  if (req.method !== "POST") {
    res.status(405).json({
      error: 'Use POST with JSON body { "question": "..." }',
    });
    return;
  }

  try {
    const { question } = req.body || {};
    const q =
      typeof question === "string" && question.trim()
        ? question.trim()
        : "What are the top 3 operational issues I should care about today, and what’s the business impact?";

    // Fetch all systems in parallel
    const [sn, sf, sp] = await Promise.all([
      fetchServiceNowSummary(),
      fetchSalesforceSummary(),
      fetchSharePointSummary(q),
    ]);

    const combinedAnswer = buildLeadershipSummary(q, sn, sf, sp);

    res.status(200).json({
      question: q,
      combinedAnswer,
      sources: {
        serviceNow: sn,
        salesforce: sf,
        sharePoint: sp,
      },
      generatedAt: new Date().toISOString(),
    });
  } catch (err) {
    console.error("[/api/txi-dashboard ERROR]:", err);
    res.status(500).json({
      error: "Internal error in /api/txi-dashboard.",
      details: String(err),
    });
  }
}
