// api/txi-dashboard.js
//
// Combined leadership view across ServiceNow, Salesforce and SharePoint.
// This version is defensive: it never throws on integration/model errors.
// Instead it returns structured `sources` + an optional `geminiError` and a
// human-readable `combinedAnswer`.
//
// Required ENV VARS:
//   SN_BASE_URL       e.g. https://ven06080.service-now.com
//   SN_USER           ServiceNow API user
//   SN_PASS           ServiceNow API password
//   SF_USERNAME
//   SF_PASSWORD
//   SF_TOKEN
//   SF_LOGIN_URL      e.g. https://login.salesforce.com
//   SP_CHAT_URL       URL of your SharePoint assistant endpoint
//   GEMINI_API_KEY
//   GEMINI_MODEL      e.g. gemini-1.5-flash (default if missing)
//
// NOTE: 429 from Gemini (quota) will *not* break the endpoint –
//       you’ll just see `geminiError` in the response.

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
  SP_CHAT_URL,
  GEMINI_API_KEY,
  GEMINI_MODEL,
} = process.env;

// ---------- Helpers: ServiceNow, Salesforce, SharePoint ----------

async function fetchServiceNowSummary() {
  if (!SN_BASE_URL || !SN_USER || !SN_PASS) {
    return {
      source: "ServiceNow",
      error: "Missing SN_BASE_URL / SN_USER / SN_PASS env vars",
    };
  }

  const url = `${SN_BASE_URL.replace(
    /\/$/,
    ""
  )}/api/dtp/schindler_txi/incident_summary`;

  try {
    const res = await fetch(url, {
      method: "GET",
      headers: {
        Authorization:
          "Basic " +
          Buffer.from(`${SN_USER}:${SN_PASS}`, "utf8").toString("base64"),
        Accept: "application/json",
      },
    });

    if (!res.ok) {
      const text = await res.text();
      return {
        source: "ServiceNow",
        error: `HTTP ${res.status}: ${text}`,
      };
    }

    const data = await res.json();
    // Expecting: { source, generatedAt, totalHighPriority, byPriority, ebcIncidents }
    return { source: "ServiceNow", ...data };
  } catch (err) {
    return {
      source: "ServiceNow",
      error: `ServiceNow fetch failed: ${err.message}`,
    };
  }
}

async function fetchSalesforceSummary() {
  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN) {
    return {
      source: "Salesforce",
      error: "Missing SF_USERNAME / SF_PASSWORD / SF_TOKEN env vars",
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
      error: `Salesforce login failed: ${err.message}`,
    };
  }

  try {
    // 1) Find one “EBC” account – use standard fields only (no custom-field crashes).
    const ebcAccount = await conn.sobject("Account").findOne(
      {
        // You *can* swap this to your custom checkbox later:
        // Is_EBC_Account__c: true
        Name: "EBC HQ",
      },
      ["Id", "Name", "Industry", "Rating"]
    );

    if (!ebcAccount) {
      return {
        source: "Salesforce",
        ebcAccount: null,
        atRiskSummary: null,
        atRiskOpportunities: [],
        note: "No EBC account found using Name = 'EBC HQ'.",
      };
    }

    // 2) Fetch open opportunities for that account using *only standard* fields.
    const opps = await conn.query(
      `
      SELECT Id, Name, Amount, StageName, CloseDate, Probability, IsClosed
      FROM Opportunity
      WHERE AccountId = '${ebcAccount.Id}'
      `
    );

    const allOpps = opps.records || [];

    // Simple “at risk” rule: open + low probability or early stage.
    const atRiskOpps = allOpps.filter((o) => {
      const prob = Number(o.Probability) || 0;
      const stage = (o.StageName || "").toLowerCase();
      const isClosed = !!o.IsClosed;

      if (isClosed) return false;
      if (prob <= 40) return true;
      if (
        stage.includes("prospecting") ||
        stage.includes("qualification") ||
        stage.includes("proposal")
      ) {
        return true;
      }
      return false;
    });

    const totalAmount = atRiskOpps.reduce(
      (sum, o) => sum + (Number(o.Amount) || 0),
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
        opportunityCount: atRiskOpps.length,
        totalAmount,
      },
      atRiskOpportunities: atRiskOpps.map((o) => ({
        id: o.Id,
        name: o.Name,
        amount: Number(o.Amount) || 0,
        stage: o.StageName,
        closeDate: o.CloseDate,
        probability: Number(o.Probability) || 0,
      })),
    };
  } catch (err) {
    // If you mess up a SOQL query / field name, we **contain** the blast radius here.
    return {
      source: "Salesforce",
      error: `Salesforce query failed: ${err.message}`,
    };
  }
}

async function fetchSharePointSummary(question) {
  if (!SP_CHAT_URL) {
    return {
      source: "SharePoint",
      error: "SP_CHAT_URL env var not configured.",
    };
  }

  try {
    const res = await fetch(SP_CHAT_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ question }),
    });

    if (!res.ok) {
      const text = await res.text();
      return {
        source: "SharePoint",
        error: `HTTP ${res.status}: ${text}`,
      };
    }

    const data = await res.json();
    // Expecting: { answer, usedFiles, candidateFiles, ... }
    return { source: "SharePoint", ...data };
  } catch (err) {
    return {
      source: "SharePoint",
      error: `SharePoint fetch failed: ${err.message}`,
    };
  }
}

// ---------- Helper: Gemini call + fallback summarisation ----------

async function callGemini(question, sources) {
  // If no key, just return a simple stitched answer.
  if (!GEMINI_API_KEY) {
    return {
      answer:
        "Gemini is not configured (missing GEMINI_API_KEY). " +
        "Here is a raw summary of the sources:\n\n" +
        JSON.stringify(sources, null, 2),
      error: "GEMINI_API_KEY not set",
    };
  }

  const model = GEMINI_MODEL || "gemini-1.5-flash";
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GEMINI_API_KEY}`;

  const prompt = `
You are an executive advisor creating a clear, sharp leadership summary.

The user asked:
"${question}"

You have three data sources:

1) Salesforce (customer & opportunities):
${JSON.stringify(sources.salesforce)}

2) ServiceNow (incident / IT health):
${JSON.stringify(sources.serviceNow)}

3) SharePoint (knowledge documents, if any):
${JSON.stringify(sources.sharePoint)}

TASK:
- Give a concise 3–5 bullet leadership view.
- Group bullets under **Sales / Revenue**, **IT / Operations**, **Collaboration / Knowledge**.
- Be specific on numbers (deal values, counts of incidents, etc.).
- End with a short 2–3 line "So what should leadership do next?" section.

Keep it simple, non-technical, and focused on impact.
`;

  try {
    const res = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        contents: [
          {
            parts: [{ text: prompt }],
          },
        ],
      }),
    });

    if (!res.ok) {
      const text = await res.text();
      return {
        answer:
          "I could not generate an AI-composed answer because Gemini returned an error. " +
          "Here is the raw data from each system:\n\n" +
          JSON.stringify(sources, null, 2),
        error: `Gemini HTTP ${res.status}: ${text}`,
      };
    }

    const data = await res.json();
    const content =
      data.candidates?.[0]?.content?.parts?.[0]?.text ||
      "No content returned by Gemini. Here is the raw data:\n\n" +
        JSON.stringify(sources, null, 2);

    return {
      answer: content,
      error: null,
    };
  } catch (err) {
    return {
      answer:
        "Gemini call failed. Here is the raw data from the systems instead:\n\n" +
        JSON.stringify(sources, null, 2),
      error: `Gemini exception: ${err.message}`,
    };
  }
}

// ---------- Main handler ----------

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res
      .status(405)
      .json({ error: 'Use POST with JSON body { "question": "..." }' });
  }

  let question;
  try {
    question = req.body?.question;
  } catch {
    // In case body parsing fails on some environments
    try {
      const text = await new Promise((resolve, reject) => {
        let data = "";
        req.on("data", (chunk) => (data += chunk));
        req.on("end", () => resolve(data));
        req.on("error", reject);
      });
      const parsed = JSON.parse(text || "{}");
      question = parsed.question;
    } catch (err) {
      return res.status(400).json({
        error: 'Invalid JSON body. Expected { "question": "..." }',
      });
    }
  }

  if (!question || typeof question !== "string") {
    return res.status(400).json({
      error: 'Missing "question" in request body or not a string.',
    });
  }

  try {
    // Call all sources in parallel but isolate failures inside helpers.
    const [serviceNow, salesforce, sharePoint] = await Promise.all([
      fetchServiceNowSummary(),
      fetchSalesforceSummary(),
      fetchSharePointSummary(question),
    ]);

    const sources = { serviceNow, salesforce, sharePoint };

    const geminiResult = await callGemini(question, sources);

    return res.status(200).json({
      question,
      combinedAnswer: geminiResult.answer,
      geminiError: geminiResult.error,
      sources,
      generatedAt: new Date().toISOString(),
    });
  } catch (err) {
    // Absolute last-resort catch – should almost never hit now.
    return res.status(500).json({
      error: "Unhandled error in txi-dashboard handler.",
      detail: err.message,
    });
  }
}
