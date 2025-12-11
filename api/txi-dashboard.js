// /api/txi-dashboard.js
// Schindler Total Intelligence – leadership console backend
// Combines data from ServiceNow, Salesforce and SharePoint,
// then asks Gemini for an executive-level, action-oriented summary.

import fetch from "node-fetch";
import jsforce from "jsforce";

const {
  // ServiceNow
  SN_TXI_URL,           // e.g. https://ven06080.service-now.com/api/dtp/schindler_txi/incident_summary
  SN_USERNAME,
  SN_PASSWORD,

  // Salesforce
  SF_USERNAME,
  SF_PASSWORD,
  SF_TOKEN,
  SF_LOGIN_URL,

  // SharePoint chat endpoint (our /api/chat-sp or a similar proxy)
  SP_CHAT_URL,

  // Gemini
  GEMINI_API_KEY,
  GEMINI_MODEL,         // optional: e.g. "gemini-2.0-flash-exp" or future "gemini-2.5" / "gemini-3.0"
} = process.env;

const DEFAULT_GEMINI_MODEL = GEMINI_MODEL || "gemini-2.0-flash-exp";

function basicAuthHeader(user, pass) {
  const token = Buffer.from(`${user}:${pass}`).toString("base64");
  return `Basic ${token}`;
}

// ---------- ServiceNow ----------

async function fetchServiceNowSummary() {
  if (!SN_TXI_URL || !SN_USERNAME || !SN_PASSWORD) {
    return {
      source: "ServiceNow",
      error: "ServiceNow env vars (SN_TXI_URL / SN_USERNAME / SN_PASSWORD) not configured.",
    };
  }

  try {
    const resp = await fetch(SN_TXI_URL, {
      method: "GET",
      headers: {
        Authorization: basicAuthHeader(SN_USERNAME, SN_PASSWORD),
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
    // Expecting your custom summary JSON from /api/dtp/schindler_txi/incident_summary
    return {
      source: "ServiceNow",
      ...data,
    };
  } catch (err) {
    return {
      source: "ServiceNow",
      error: `ServiceNow fetch failed: ${err.message}`,
    };
  }
}

// ---------- Salesforce ----------

async function fetchSalesforceSummary() {
  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN || !SF_LOGIN_URL) {
    return {
      source: "Salesforce",
      error: "Salesforce env vars not configured (SF_USERNAME / SF_PASSWORD / SF_TOKEN / SF_LOGIN_URL).",
    };
  }

  const conn = new jsforce.Connection({ loginUrl: SF_LOGIN_URL });

  try {
    await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);

    // 1) Find the EBC account (POC: we hard-code the sample account by name).
    const ebcAccount = await conn
      .sobject("Account")
      .findOne(
        { Name: "EBC HQ" },
        "Id, Name, Industry, Rating, CustomerPriority"
      );

    if (!ebcAccount) {
      return {
        source: "Salesforce",
        error: "No EBC HQ account found in Salesforce (expected Name = 'EBC HQ').",
      };
    }

    // 2) Get open opportunities for that account.
    //    For the POC we treat *all* open opps for EBC HQ as “at-risk sample deals”.
    const oppResult = await conn.query(
      `
      SELECT Id, Name, Amount, StageName, CloseDate, Probability
      FROM Opportunity
      WHERE AccountId = '${ebcAccount.Id}'
        AND IsClosed = false
      ORDER BY CloseDate ASC
      `
    );

    const atRiskOpportunities = oppResult.records.map((o) => ({
      id: o.Id,
      name: o.Name,
      amount: o.Amount,
      stage: o.StageName,
      closeDate: o.CloseDate,
      probability: o.Probability,
    }));

    const atRiskSummary = {
      opportunityCount: atRiskOpportunities.length,
      totalAmount: atRiskOpportunities.reduce(
        (sum, o) => sum + (o.amount || 0),
        0
      ),
    };

    return {
      source: "Salesforce",
      ebcAccount: {
        id: ebcAccount.Id,
        name: ebcAccount.Name,
        industry: ebcAccount.Industry,
        rating: ebcAccount.Rating,
        customerPriority: ebcAccount.CustomerPriority,
      },
      atRiskSummary,
      atRiskOpportunities,
    };
  } catch (err) {
    return {
      source: "Salesforce",
      error: `Salesforce error: ${err.message}`,
    };
  }
}

// ---------- SharePoint ----------

async function fetchSharePointSummary(question) {
  if (!SP_CHAT_URL) {
    return {
      source: "SharePoint",
      error: "SP_CHAT_URL env var not configured.",
    };
  }

  // We *explicitly* nudge the SharePoint endpoint towards the three sample docs
  // you created: Annual EBC Review Notes, EBC_Account_Health_Risk, IT_Operations_Weekly_Report.
  const spQuestion = `
${question}

You are preparing a Schindler leadership dashboard.
FOCUS FIRST on any documents whose name or contents match:

- "Annual EBC Review Notes"
- "EBC_Account_Health_Risk"
- "IT_Operations_Weekly_Report"

Use them as primary sources to extract:
- EBC account health & relationship risks
- Customer sentiment / escalation signals
- IT operations stability, outages, and repetitive issues
- Any explicit recommendations already written in those docs.
`;

  try {
    const resp = await fetch(SP_CHAT_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ question: spQuestion }),
    });

    if (!resp.ok) {
      const text = await resp.text();
      return {
        source: "SharePoint",
        error: `SharePoint chat endpoint error ${resp.status}: ${text}`,
      };
    }

    const data = await resp.json();
    // Expecting shape: { answer, usedFiles, candidateFiles, ... }
    return {
      source: "SharePoint",
      ...data,
    };
  } catch (err) {
    return {
      source: "SharePoint",
      error: `SharePoint fetch failed: ${err.message}`,
    };
  }
}

// ---------- Gemini ----------

async function callGemini(question, serviceNow, salesforce, sharePoint) {
  if (!GEMINI_API_KEY) {
    return {
      text:
        "Gemini API key not configured. Cannot generate combined leadership view.",
    };
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${DEFAULT_GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;

  const payload = {
    contents: [
      {
        role: "user",
        parts: [
          {
            text: `
You are the AI brain behind "Schindler Total Intelligence – Leadership Console".

You receive structured JSON from three systems:

1) "serviceNow": IT incident summary for today.
2) "salesforce": key account information and open opportunities for EBC HQ.
3) "sharePoint": summaries of leadership documents plus the list of files used.

User question:
"${question}"

Your job:
- Produce a *clear, executive-level briefing* that a CEO or BU head can read in under 1 minute.
- Focus on *today’s* biggest risks and customer impacts.
- Use simple language, no technical jargon.

Structure your answer exactly like this:

1. A short opening paragraph (2–3 sentences) giving the overall picture.
2. Then 3–5 numbered bullets. For each bullet:
   - A title in **bold** (one short phrase).
   - 2–3 sentences answering:
       • What is happening?  
       • Why does it matter in business terms (revenue, CX, EX, risk)?
       • What should leadership do in the next 24–72 hours?
3. A final single-sentence "Bottom line" starting with "Bottom line:".

Be very explicit and actionable. If some source returns an error, call it out as "data visibility risk" instead of pretending everything is fine.

Here is the raw JSON context:

SERVICE_NOW:
${JSON.stringify(serviceNow, null, 2)}

SALESFORCE:
${JSON.stringify(salesforce, null, 2)}

SHAREPOINT:
${JSON.stringify(sharePoint, null, 2)}
          `,
          },
        ],
      },
    ],
  };

  const resp = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  if (!resp.ok) {
    const text = await resp.text();
    return {
      text: `Gemini error ${resp.status}: ${text}`,
    };
  }

  const data = await resp.json();

  const text =
    data.candidates?.[0]?.content?.parts?.[0]?.text ||
    "Gemini did not return any text.";

  return { text, raw: data };
}

// ---------- Main handler ----------

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({
      error: 'Use POST with JSON body { "question": "..." }',
    });
  }

  let question;

  try {
    ({ question } = req.body || {});
  } catch (err) {
    return res.status(400).json({
      error: "Invalid JSON in request body.",
    });
  }

  if (!question || typeof question !== "string") {
    return res.status(400).json({
      error: 'Missing "question" in request body or not a string.',
    });
  }

  try {
    // Run all 3 systems in parallel.
    const [serviceNow, salesforce, sharePoint] = await Promise.all([
      fetchServiceNowSummary(),
      fetchSalesforceSummary(),
      fetchSharePointSummary(question),
    ]);

    const combined = await callGemini(
      question,
      serviceNow,
      salesforce,
      sharePoint
    );

    return res.status(200).json({
      question,
      combinedAnswer: combined.text,
      sources: {
        serviceNow,
        salesforce,
        sharePoint,
      },
      generatedAt: new Date().toISOString(),
    });
  } catch (err) {
    console.error("TXI dashboard fatal error:", err);
    return res.status(500).json({
      error: "TXI dashboard internal error.",
      detail: err.message,
    });
  }
}
