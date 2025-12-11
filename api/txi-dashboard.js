// api/txi-dashboard.js
//
// Schindler Total Intelligence – TXI Dashboard API
// Combines:
//  - ServiceNow incident summary (REST API)
//  - Salesforce EBC account + at-risk opps (jsforce)
//  - SharePoint doc summary (chat-sp endpoint)
//
// Then:
//  - Tries Gemini (model from GEMINI_MODEL) to compose a leadership answer
//  - If Gemini fails or is disabled, falls back to a local leadership summary
//
// This MUST NOT throw for normal data issues – always return 200 with something usable.

import fetch from "node-fetch";
import jsforce from "jsforce";

// ---------- ENVIRONMENT VARIABLES ----------

const {
  // ServiceNow
  SN_TXI_URL,          // e.g. https://ven06080.service-now.com/api/dtp/schindler_txi/incident_summary
  SN_USERNAME,
  SN_PASSWORD,

  // Salesforce
  SF_USERNAME,
  SF_PASSWORD,
  SF_TOKEN,
  SF_LOGIN_URL,        // e.g. https://login.salesforce.com or https://test.salesforce.com

  // SharePoint chat endpoint
  SP_CHAT_URL,         // e.g. https://<your-vercel-app>.vercel.app/api/chat-sp

  // Gemini
  GEMINI_API_KEY,
  GEMINI_MODEL         // e.g. gemini-2.0-flash, gemini-1.5-flash, etc.
} = process.env;

// ---------- HELPERS ----------

function basicAuthHeader(user, pass) {
  const token = Buffer.from(`${user}:${pass}`).toString("base64");
  return `Basic ${token}`;
}

// ---------- SERVICE NOW ----------

async function fetchServiceNowSummary() {
  if (!SN_TXI_URL || !SN_USERNAME || !SN_PASSWORD) {
    return {
      source: "ServiceNow",
      error: "ServiceNow env vars not configured (SN_TXI_URL / SN_USERNAME / SN_PASSWORD)."
    };
  }

  try {
    const resp = await fetch(SN_TXI_URL, {
      method: "GET",
      headers: {
        "Authorization": basicAuthHeader(SN_USERNAME, SN_PASSWORD),
        "Accept": "application/json"
      }
    });

    const text = await resp.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch (e) {
      return {
        source: "ServiceNow",
        error: "ServiceNow returned non-JSON response."
      };
    }

    if (!resp.ok) {
      return {
        source: "ServiceNow",
        error: data.error || `ServiceNow API error (status ${resp.status}).`
      };
    }

    // Expecting summary-style payload:
    // {
    //   source: "ServiceNow",
    //   generatedAt: "...",
    //   totalHighPriority: number,
    //   byPriority: [{priority:"1",count:72},...],
    //   ebcIncidents: [...]
    // }
    return {
      source: "ServiceNow",
      generatedAt: data.generatedAt || null,
      totalHighPriority: data.totalHighPriority ?? null,
      byPriority: Array.isArray(data.byPriority) ? data.byPriority : [],
      ebcIncidents: Array.isArray(data.ebcIncidents) ? data.ebcIncidents : []
    };
  } catch (err) {
    console.error("[TXI] ServiceNow error (masked):", String(err));
    return {
      source: "ServiceNow",
      error: "Error calling ServiceNow API."
    };
  }
}

// ---------- SALESFORCE ----------

async function fetchSalesforceSummary() {
  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN || !SF_LOGIN_URL) {
    return {
      source: "Salesforce",
      error: "Salesforce env vars not fully configured (SF_USERNAME / SF_PASSWORD / SF_TOKEN / SF_LOGIN_URL)."
    };
  }

  const conn = new jsforce.Connection({
    loginUrl: SF_LOGIN_URL
  });

  try {
    await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);

    // 1) Get EBC HQ account (simple Name-based for POC)
    const accounts = await conn.query(
      "SELECT Id, Name, Industry, Rating FROM Account WHERE Name = 'EBC HQ' LIMIT 1"
    );
    const ebcAccount = (accounts.records && accounts.records[0]) || null;

    // 2) Get at-risk Opps for that account (simple rule: Probability <= 40)
    let atRiskSummary = null;
    let atRiskOpportunities = [];

    if (ebcAccount) {
      const oppsRes = await conn.query(
        `SELECT Id, Name, Amount, StageName, CloseDate, Probability
         FROM Opportunity
         WHERE AccountId = '${ebcAccount.Id}'
         LIMIT 50`
      );

      const allOpps = oppsRes.records || [];
      atRiskOpportunities = allOpps.filter((o) => {
        const prob = typeof o.Probability === "number" ? o.Probability : 0;
        return prob <= 40;
      });

      const totalAmount = atRiskOpportunities.reduce(
        (sum, o) => sum + (o.Amount || 0),
        0
      );

      atRiskSummary = {
        opportunityCount: atRiskOpportunities.length,
        totalAmount
      };
    }

    return {
      source: "Salesforce",
      ebcAccount: ebcAccount
        ? {
            id: ebcAccount.Id,
            name: ebcAccount.Name,
            industry: ebcAccount.Industry || null,
            rating: ebcAccount.Rating || null
          }
        : null,
      atRiskSummary,
      atRiskOpportunities: atRiskOpportunities.map((o) => ({
        id: o.Id,
        name: o.Name,
        amount: o.Amount || 0,
        stage: o.StageName,
        closeDate: o.CloseDate,
        probability: o.Probability
      }))
    };
  } catch (err) {
    console.error("[TXI] Salesforce error (masked):", String(err));
    return {
      source: "Salesforce",
      error: String(err)
    };
  }
}

// ---------- SHAREPOINT ----------

async function fetchSharePointSummary(question) {
  if (!SP_CHAT_URL) {
    return {
      source: "SharePoint",
      error: "SP_CHAT_URL env var not configured."
    };
  }

  try {
    const resp = await fetch(SP_CHAT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        // For leadership TXI we can ignore their wording and just ask
        // the SP assistant to surface TXI / risk / execution docs.
        question:
          question ||
          'Find and summarise documents related to deals, execution plans, and risk.'
      })
    });

    const text = await resp.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch (e) {
      return {
        source: "SharePoint",
        error: "SharePoint /api/chat-sp returned non-JSON."
      };
    }

    // Expected from chat-sp:
    // { answer: "...", usedFiles: [...], candidateFiles: [...] }
    if (!resp.ok) {
      return {
        source: "SharePoint",
        error: data.error || `SharePoint chat-sp error (status ${resp.status}).`
      };
    }

    return {
      source: "SharePoint",
      answer: data.answer || "",
      usedFiles: Array.isArray(data.usedFiles) ? data.usedFiles : [],
      candidateFiles: Array.isArray(data.candidateFiles) ? data.candidateFiles : []
    };
  } catch (err) {
    console.error("[TXI] SharePoint error (masked):", String(err));
    return {
      source: "SharePoint",
      error: "Error calling SharePoint chat-sp endpoint."
    };
  }
}

// ---------- GEMINI COMPOSER (OPTIONAL) ----------

async function callGeminiComposer(question, sources) {
  if (!GEMINI_API_KEY || !GEMINI_MODEL) {
    return null; // Gemini not configured → skip
  }

  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;

    const prompt = `
You are an executive assistant for Schindler leadership.

The user asked:
${question}

You have three JSON blocks with live signals from different systems:

1) ServiceNow (IT incidents and high-priority load)
${JSON.stringify(sources.serviceNow, null, 2)}

2) Salesforce (EBC HQ account and at-risk opportunities)
${JSON.stringify(sources.salesforce, null, 2)}

3) SharePoint (document insights)
${JSON.stringify(sources.sharePoint, null, 2)}

TASK:
- Give a SINGLE leadership summary of today's top operational issues and business impacts.
- Use clear headings and bullet points.
- Prioritise:
  1) Revenue / customer risk from Salesforce
  2) IT stability / outages from ServiceNow
  3) Knowledge / collaboration gaps or insights from SharePoint
- Be concise but executive-level (no raw JSON, no technical noise).
- If some sources have errors or no data, mention that clearly but briefly.
- Do NOT invent numbers or systems that aren't in the JSON.
`.trim();

    const body = {
      contents: [
        {
          parts: [{ text: prompt }]
        }
      ]
    };

    const resp = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });

    const raw = await resp.text();
    let data;
    try {
      data = JSON.parse(raw);
    } catch (e) {
      console.error("[TXI] Gemini non-JSON response (masked).");
      return null;
    }

    if (!resp.ok || !data.candidates || !data.candidates.length) {
      console.error("[TXI] Gemini API error (masked):", data.error || data);
      return null;
    }

    const candidate = data.candidates[0];
    const parts = candidate.content && candidate.content.parts;
    const answer = parts && parts[0] && parts[0].text;

    if (!answer || !answer.trim()) return null;
    return answer.trim();
  } catch (err) {
    console.error("[TXI] Gemini call failed (masked):", String(err));
    return null;
  }
}

// ---------- LOCAL FALLBACK SUMMARY (NO GEMINI) ----------

function buildFallbackLeadershipSummary(question, { serviceNow, salesforce, sharePoint }) {
  const lines = [];
  lines.push("Here's a leadership view of today's biggest issues and business impact based on the available data:\n");

  // 1) Salesforce – revenue risk
  const sf = salesforce || {};
  if (sf.error) {
    lines.push(`* **Sales / Revenue visibility issue (Salesforce)**`);
    lines.push(`  * Business impact: Salesforce reported an error: ${sf.error}. We may not have a clear view of key accounts or at-risk opportunities.\n`);
  } else if (sf.ebcAccount && sf.atRiskSummary && sf.atRiskSummary.opportunityCount > 0) {
    lines.push(`* **Revenue risk on key customer account (Salesforce)**`);
    lines.push(
      `  * Account: ${sf.ebcAccount.name} (Industry: ${sf.ebcAccount.industry || "N/A"}, Rating: ${sf.ebcAccount.rating || "N/A"})`
    );
    lines.push(
      `  * Open at-risk opportunities: ${sf.atRiskSummary.opportunityCount} deal(s) with total value around $${sf.atRiskSummary.totalAmount}`
    );
    lines.push(
      `  * Leadership impact: Losing or delaying these deals directly impacts revenue and the relationship with this strategic account.\n`
    );
  } else {
    lines.push(`* **No explicit at-risk revenue identified (Salesforce)**`);
    lines.push(
      `  * Business impact: Current sample data does not highlight specific at-risk deals, but we should confirm whether risk-flagging and pipeline health logic are correctly configured.`
    );
    lines.push("");
  }

  // 2) ServiceNow – high priority incidents
  const sn = serviceNow || {};
  if (sn.error) {
    lines.push(`* **IT service visibility gap (ServiceNow)**`);
    lines.push(
      `  * Business impact: ServiceNow data could not be retrieved ("${sn.error}"). Leadership lacks visibility into high-priority incidents, which can hide outages and user-impacting issues.`
    );
    lines.push("");
  } else if (typeof sn.totalHighPriority === "number") {
    const list = Array.isArray(sn.byPriority) ? sn.byPriority : [];
    const p1 = list.find((p) => p.priority === "1");
    const p2 = list.find((p) => p.priority === "2");
    const p3 = list.find((p) => p.priority === "3");

    lines.push(`* **High-priority IT incident load (ServiceNow)**`);
    lines.push(`  * High-priority incidents open: ${sn.totalHighPriority}`);
    lines.push(
      `  * Distribution by priority: P1: ${p1 ? p1.count : 0}, P2: ${p2 ? p2.count : 0}, P3: ${p3 ? p3.count : 0}`
    );
    lines.push(
      `  * Leadership impact: Persistent high-priority incidents put uptime, employee productivity and customer experience at risk.`
    );
    lines.push("");
  } else {
    lines.push(`* **No high-priority incident summary available (ServiceNow)**`);
    lines.push(`  * Business impact: Without a simple summary of high-priority incidents, leadership cannot quickly gauge IT stability.`);
    lines.push("");
  }

  // 3) SharePoint – collaboration / knowledge
  const sp = sharePoint || {};
  if (sp.error) {
    lines.push(`* **Collaboration / knowledge visibility gap (SharePoint)**`);
    lines.push(`  * Business impact: ${sp.error}`);
    lines.push(
      `  * This prevents us from turning key documents (execution plans, risk registers, customer decks) into a simple leadership signal.`
    );
    lines.push("");
  } else if (sp.answer) {
    lines.push(`* **Knowledge insights from SharePoint (documents)**`);
    lines.push(`  * Assistant summary: ${sp.answer}`);
    lines.push("");
  } else {
    lines.push(`* **No document insights returned from SharePoint**`);
    lines.push(
      `  * Business impact: Without summarised knowledge from key documents, decision-making and alignment rely heavily on manual digging.`
    );
    lines.push("");
  }

  lines.push("---");
  lines.push(
    "This is a Proof of Concept using limited sample data, but it shows how a single question can pull together revenue risk, IT stability, and collaboration signals into one executive view."
  );

  return lines.join("\n");
}

// ---------- MAIN HANDLER ----------

export default async function handler(req, res) {
  if (req.method === "OPTIONS") {
    res.status(200).end();
    return;
  }

  if (req.method !== "POST") {
    res
      .status(405)
      .json({ error: 'Use POST with JSON body { "question": "..." } on /api/txi-dashboard.' });
    return;
  }

  try {
    const { question } = req.body || {};
    const q =
      typeof question === "string" && question.trim().length
        ? question.trim()
        : "What are the top 3 operational issues I should care about today, and what’s the business impact?";

    // Fetch all three systems in parallel
    const [serviceNow, salesforce, sharePoint] = await Promise.all([
      fetchServiceNowSummary(),
      fetchSalesforceSummary(),
      fetchSharePointSummary(q)
    ]);

    const sources = { serviceNow, salesforce, sharePoint };

    // Try Gemini first (if configured)
    let combinedAnswer = null;
    let usedGemini = false;

    const geminiAnswer = await callGeminiComposer(q, sources);
    if (geminiAnswer) {
      combinedAnswer = geminiAnswer;
      usedGemini = true;
    } else {
      // Fallback – local summary, no Gemini
      combinedAnswer = buildFallbackLeadershipSummary(q, sources);
    }

    res.status(200).json({
      question: q,
      combinedAnswer,
      sources,
      usedGemini,
      generatedAt: new Date().toISOString()
    });
  } catch (err) {
    console.error("[TXI] Fatal handler error (masked):", String(err));
    // LAST RESORT – return 500, but this should be rare now.
    res.status(500).json({
      error: "Internal error in /api/txi-dashboard.",
      details: "Check Vercel logs for [TXI] errors."
    });
  }
}
