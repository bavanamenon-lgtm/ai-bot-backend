// /api/txi-dashboard.js
// Combined leadership view across ServiceNow, Salesforce and SharePoint.

import fetch from "node-fetch";
import jsforce from "jsforce";

const {
  SN_BASE_URL,
  SN_USER,
  SN_PASS,
  SN_API_PATH,          // e.g. "/api/dtp/schindler_txi/incident_summary"
  SF_USERNAME,
  SF_PASSWORD,
  SF_TOKEN,
  SF_LOGIN_URL,
  GEMINI_API_KEY,
  SP_CHAT_SP_URL        // e.g. "https://ai-bot-backend-black.vercel.app/api/chat-sp"
} = process.env;

// ---------- ServiceNow ----------

async function fetchServiceNowIncidents() {
  if (!SN_BASE_URL || !SN_USER || !SN_PASS || !SN_API_PATH) {
    console.warn("[TXI] ServiceNow env vars missing – skipping");
    return { error: "ServiceNow not configured", source: "ServiceNow" };
  }

  const url = `${SN_BASE_URL}${SN_API_PATH}`;

  try {
    const res = await fetch(url, {
      method: "GET",
      headers: {
        Accept: "application/json",
        Authorization:
          "Basic " +
          Buffer.from(`${SN_USER}:${SN_PASS}`, "utf8").toString("base64"),
      },
    });

    if (!res.ok) {
      console.error("[TXI] ServiceNow HTTP error", res.status);
      return { error: "Error: ServiceNow API error", source: "ServiceNow" };
    }

    const data = await res.json();

    // Expecting something like: { totals: {...}, topPriorities: [...] }
    return {
      source: "ServiceNow",
      raw: data,
    };
  } catch (err) {
    console.error("[TXI] ServiceNow fetch failed", err);
    return { error: "Error: ServiceNow API error", source: "ServiceNow" };
  }
}

// ---------- Salesforce ----------

async function fetchSalesforceView() {
  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN || !SF_LOGIN_URL) {
    console.warn("[TXI] Salesforce env vars missing – skipping");
    return { source: "Salesforce", ebcAccount: null, atRiskOpportunities: [] };
  }

  const conn = new jsforce.Connection({ loginUrl: SF_LOGIN_URL });

  try {
    await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);

    // IMPORTANT: adjust these three field API names if your org uses different ones.
    const EBC_FIELD = "Is_EBC__c";
    const RISK_FLAG_FIELD = "Risk_Flag__c";
    const AT_RISK_FIELD = "At_Risk__c";

    // 1) Find one EBC account
    const ebcQuery = `
      SELECT Id, Name, Industry, CustomerPriority__c, Active__c
      FROM Account
      WHERE ${EBC_FIELD} = true
      LIMIT 1
    `;

    const ebcResult = await conn.query(ebcQuery);
    const ebcAccount = ebcResult.records?.[0] || null;

    // 2) Find at-risk opportunities (linked to EBC account if present)
    let oppWhere = `(${RISK_FLAG_FIELD} = true OR ${AT_RISK_FIELD} = true)`;
    if (ebcAccount) {
      oppWhere += ` AND AccountId = '${ebcAccount.Id}'`;
    }

    const oppQuery = `
      SELECT Id, Name, Amount, StageName, CloseDate,
             ${RISK_FLAG_FIELD}, ${AT_RISK_FIELD},
             AccountId
      FROM Opportunity
      WHERE ${oppWhere}
      ORDER BY CloseDate ASC
      LIMIT 5
    `;

    const oppResult = await conn.query(oppQuery);
    const atRiskOpportunities = oppResult.records || [];

    console.log(
      "[TXI] SF debug",
      JSON.stringify({
        ebcFound: !!ebcAccount,
        ebcName: ebcAccount?.Name,
        oppCount: atRiskOpportunities.length,
      })
    );

    return {
      source: "Salesforce",
      ebcAccount,
      atRiskOpportunities,
    };
  } catch (err) {
    console.error("[TXI] Salesforce error", err);
    return {
      source: "Salesforce",
      ebcAccount: null,
      atRiskOpportunities: [],
      error: "Salesforce query error",
    };
  }
}

// ---------- SharePoint (existing /api/chat-sp) ----------

async function fetchSharePointSummary(question) {
  if (!SP_CHAT_SP_URL) {
    console.warn("[TXI] SharePoint env var SP_CHAT_SP_URL missing – skipping");
    return { source: "SharePoint", error: "SharePoint not configured" };
  }

  try {
    const res = await fetch(SP_CHAT_SP_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ question }),
    });

    if (!res.ok) {
      console.error("[TXI] SharePoint API non-200", res.status);
      return {
        source: "SharePoint",
        error: "Error: SharePoint /api/chat-sp returned non-JSON",
      };
    }

    const data = await res.json();
    return {
      source: "SharePoint",
      raw: data,
    };
  } catch (err) {
    console.error("[TXI] SharePoint fetch failed", err);
    return {
      source: "SharePoint",
      error: "Error: SharePoint /api/chat-sp returned non-JSON",
    };
  }
}

// ---------- Gemini ----------

async function callGemini(question, sources) {
  if (!GEMINI_API_KEY) {
    console.warn("[TXI] GEMINI_API_KEY missing – returning fallback text");
    return (
      "AI key not configured. This is a placeholder answer for the leadership view."
    );
  }

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent";

  const payload = {
    contents: [
      {
        parts: [
          {
            text: `
You are an executive assistant.

Question from leadership:
"${question}"

ServiceNow data (maybe error):
${JSON.stringify(sources.serviceNow)}

Salesforce data:
${JSON.stringify(sources.salesforce)}

SharePoint data:
${JSON.stringify(sources.sharePoint)}

1. Give a short, leadership-friendly answer.
2. List the top 3 operational issues and the business impact.
3. Be explicit about which system each issue comes from.
4. If a system returned an error, treat that as a visibility/risk issue.
`,
          },
        ],
      },
    ],
  };

  const res = await fetch(`${url}?key=${GEMINI_API_KEY}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    console.error("[TXI] Gemini HTTP error", res.status);
    return "Could not generate an answer.";
  }

  const data = await res.json();
  const text =
    data?.candidates?.[0]?.content?.parts?.[0]?.text ||
    "Could not generate an answer.";
  return text;
}

// ---------- HTTP handler ----------

export default async function handler(req, res) {
  if (req.method !== "GET") {
    return res
      .status(405)
      .json({ error: 'Use GET with query param ?question="..."' });
  }

  const question = (req.query.question || "").toString().trim();
  if (!question) {
    return res
      .status(400)
      .json({ error: 'Missing "question" query parameter.' });
  }

  try {
    const [serviceNow, salesforce, sharePoint] = await Promise.all([
      fetchServiceNowIncidents(),
      fetchSalesforceView(),
      fetchSharePointSummary(question),
    ]);

    const combinedAnswer = await callGemini(question, {
      serviceNow,
      salesforce,
      sharePoint,
    });

    res.status(200).json({
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
    console.error("[TXI] Handler error", err);
    res.status(500).json({
      error: "TXI dashboard error",
      detail: err.message,
    });
  }
}
