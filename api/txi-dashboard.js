// /api/txi-dashboard.js
// Combined leadership view across ServiceNow, Salesforce and SharePoint.

import fetch from "node-fetch";
import jsforce from "jsforce";

const {
  // ServiceNow
  SN_BASE_URL,
  SN_USER,
  SN_PASS,

  // Salesforce (same envs you already used in chat.js)
  SF_USERNAME,
  SF_PASSWORD,
  SF_TOKEN,
  SF_LOGIN_URL,

  // SharePoint – we will call the same Vercel API you test in Postman
  // e.g. SP_CHAT_URL = "https://ai-bot-backend-black.vercel.app/api/chat-sp"
  SP_CHAT_URL,

  // Gemini for summarisation (optional but keeps the story nice)
  GEMINI_API_KEY,
} = process.env;

function basicAuth(user, pass) {
  return "Basic " + Buffer.from(`${user}:${pass}`).toString("base64");
}

async function callServiceNow() {
  if (!SN_BASE_URL || !SN_USER || !SN_PASS) {
    return { error: "ServiceNow env vars missing", source: "ServiceNow" };
  }

  try {
    const url = `${SN_BASE_URL}/api/dtp/schindler_txi/incident_summary`;

    const resp = await fetch(url, {
      method: "GET",
      headers: {
        Accept: "application/json",
        Authorization: basicAuth(SN_USER, SN_PASS),
      },
    });

    if (!resp.ok) {
      const text = await resp.text();
      throw new Error(`ServiceNow HTTP ${resp.status}: ${text.slice(0, 200)}`);
    }

    const data = await resp.json();
    return { ...data, source: "ServiceNow" };
  } catch (err) {
    return { error: `ServiceNow API error: ${err.message}`, source: "ServiceNow" };
  }
}

async function callSalesforce() {
  if (!SF_USERNAME || !SF_PASSWORD || !SF_TOKEN || !SF_LOGIN_URL) {
    return {
      source: "Salesforce",
      error: "Salesforce env vars missing",
      ebcAccount: null,
      atRiskOpportunities: [],
    };
  }

  const conn = new jsforce.Connection({ loginUrl: SF_LOGIN_URL });

  try {
    await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);

    // 1) EBC account (Is_EBC_Account__c must be the API name of your checkbox)
    const ebcResult = await conn
      .sobject("Account")
      .find({ Is_EBC_Account__c: true })
      .limit(1)
      .execute();

    const ebcAccount = ebcResult && ebcResult[0] ? ebcResult[0] : null;

    // 2) At-risk opportunities for that EBC account
    let atRiskOpportunities = [];
    if (ebcAccount) {
      atRiskOpportunities = await conn
        .sobject("Opportunity")
        .find(
          {
            AccountId: ebcAccount.Id,
            Risk_Flag__c: true, // your custom fields
            At_Risk__c: true,
          },
          [
            "Id",
            "Name",
            "Amount",
            "StageName",
            "CloseDate",
            "Probability",
            "AccountId",
          ]
        )
        .execute();
    }

    return {
      source: "Salesforce",
      ebcAccount,
      atRiskOpportunities,
    };
  } catch (err) {
    return {
      source: "Salesforce",
      error: `Salesforce API error: ${err.message}`,
      ebcAccount: null,
      atRiskOpportunities: [],
    };
  }
}

async function callSharePoint(questionForDocs) {
  if (!SP_CHAT_URL) {
    return { source: "SharePoint", error: "SP_CHAT_URL env var missing" };
  }

  try {
    const resp = await fetch(SP_CHAT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        // pass the CEO question straight through
        question: questionForDocs,
      }),
    });

    if (!resp.ok) {
      const text = await resp.text();
      throw new Error(`SharePoint HTTP ${resp.status}: ${text.slice(0, 200)}`);
    }

    const data = await resp.json(); // your /api/chat-sp returns JSON
    return { ...data, source: "SharePoint" };
  } catch (err) {
    return {
      source: "SharePoint",
      error: `SharePoint /api/chat-sp error: ${err.message}`,
    };
  }
}

async function summariseWithGemini(payload) {
  if (!GEMINI_API_KEY) {
    // No Gemini? Return a plain JS summary.
    const { serviceNow, salesforce, sharePoint, question } = payload;
    return buildPlainSummary(question, serviceNow, salesforce, sharePoint);
  }

  const prompt = `
You are building a Total Intelligence executive view for Schindler.

Question from leadership:
"${payload.question}"

ServiceNow summary JSON:
${JSON.stringify(payload.serviceNow, null, 2)}

Salesforce summary JSON:
${JSON.stringify(payload.salesforce, null, 2)}

SharePoint summary JSON:
${JSON.stringify(payload.sharePoint, null, 2)}

1. In 2–3 short paragraphs, give a leadership-level answer to the question.
2. Then list exactly 3 bullet points: "Top 3 issues and impacts".
3. Use plain business language, no JSON, no technical error codes.
4. If some source has an error, mention it once under the relevant bullet, not as a separate issue.
`;

  const resp = await fetch(
    "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" +
      GEMINI_API_KEY,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
      }),
    }
  );

  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`Gemini error ${resp.status}: ${text.slice(0, 200)}`);
  }

  const data = await resp.json();
  const text =
    data?.candidates?.[0]?.content?.parts?.map((p) => p.text).join("") ??
    "Gemini did not return any text.";
  return text;
}

// Fallback summariser if you don't want Gemini
function buildPlainSummary(question, serviceNow, salesforce, sharePoint) {
  const lines = [];

  lines.push(
    `Here is a leadership view for the question: "${question}".`,
    ""
  );

  // Salesforce
  if (salesforce.error) {
    lines.push(
      `• Salesforce: ${salesforce.error}. We cannot reliably see EBC or at-risk opportunities.`
    );
  } else if (
    salesforce.ebcAccount &&
    salesforce.atRiskOpportunities &&
    salesforce.atRiskOpportunities.length > 0
  ) {
    const total = salesforce.atRiskOpportunities.reduce(
      (sum, opp) => sum + (opp.Amount || 0),
      0
    );
    lines.push(
      `• Salesforce: EBC account "${salesforce.ebcAccount.Name}" has ${salesforce.atRiskOpportunities.length} at-risk opportunities worth ~$${total.toLocaleString()}.`
    );
  } else {
    lines.push(
      `• Salesforce: No at-risk opportunities are currently flagged for the EBC account in this sample data.`
    );
  }

  // ServiceNow
  if (serviceNow.error) {
    lines.push(`• ServiceNow: ${serviceNow.error}.`);
  } else if (serviceNow.totalHighPriority != null) {
    lines.push(
      `• ServiceNow: There are ${serviceNow.totalHighPriority} high-priority incidents in the system today, across priorities: ${JSON.stringify(
        serviceNow.byPriority || []
      )}.`
    );
  } else {
    lines.push(
      `• ServiceNow: Incident summary endpoint returned data but without "totalHighPriority".`
    );
  }

  // SharePoint
  if (sharePoint.error) {
    lines.push(`• SharePoint: ${sharePoint.error}.`);
  } else if (sharePoint.answer) {
    lines.push(
      `• SharePoint: Key document insight – ${sharePoint.answer.replace(
        /\n+/g,
        " "
      )}`
    );
  } else {
    lines.push(
      `• SharePoint: The document bot returned no specific answer for this question.`
    );
  }

  return lines.join("\n");
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res
      .status(405)
      .json({ error: 'Use POST with JSON body { "question": "..." }' });
  }

  try {
    const body =
      typeof req.body === "string" ? JSON.parse(req.body || "{}") : req.body;
    const question = body.question;

    if (!question || typeof question !== "string") {
      return res.status(400).json({
        error: 'Missing "question" in request body or not a string.',
      });
    }

    // Run all three in parallel
    const [serviceNow, salesforce, sharePoint] = await Promise.all([
      callServiceNow(),
      callSalesforce(),
      callSharePoint(question),
    ]);

    let combinedAnswer;
    try {
      combinedAnswer = await summariseWithGemini({
        question,
        serviceNow,
        salesforce,
        sharePoint,
      });
    } catch (gemErr) {
      // Gemini failed -> fall back to plain summary
      combinedAnswer = buildPlainSummary(
        question,
        serviceNow,
        salesforce,
        sharePoint
      );
    }

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
    res.status(500).json({ error: err.message || "Unexpected server error" });
  }
}
