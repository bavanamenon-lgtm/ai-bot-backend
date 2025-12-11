// /api/txi-dashboard.js

import { NextResponse } from "next/server";

// ----------------------
// ENV VARIABLES REQUIRED
// ----------------------
const {
  SF_USERNAME,
  SF_PASSWORD,
  SF_TOKEN,
  SF_LOGIN_URL,
  SN_URL,
  SN_USER,
  SN_PASS,
  SP_CHAT_URL,
  GEMINI_API_KEY
} = process.env;

// ----------------------
// SALESFORCE CLIENT
// ----------------------
import jsforce from "jsforce";

// Query Account + At-Risk Opportunities
async function fetchSalesforceData() {
  try {
    const conn = new jsforce.Connection({
      loginUrl: SF_LOGIN_URL,
    });

    await conn.login(SF_USERNAME, SF_PASSWORD + SF_TOKEN);

    // 1) Get EBC Account
    const account = await conn.query(`
      SELECT Id, Name, Industry, Rating
      FROM Account
      WHERE Name = 'EBC HQ'
      LIMIT 1
    `);

    if (account.totalSize === 0) {
      return { source: "Salesforce", error: "EBC account not found." };
    }

    const acc = account.records[0];

    // 2) Fetch at-risk opportunities
    const opps = await conn.query(`
      SELECT Id, Name, Amount, StageName, CloseDate, Probability
      FROM Opportunity
      WHERE AccountId = '${acc.Id}' AND At_Risk__c = TRUE
    `);

    const totalAmount = opps.records.reduce((t, o) => t + (o.Amount || 0), 0);

    return {
      source: "Salesforce",
      ebcAccount: {
        id: acc.Id,
        name: acc.Name,
        industry: acc.Industry,
        rating: acc.Rating
      },
      atRiskSummary: {
        opportunityCount: opps.totalSize,
        totalAmount
      },
      atRiskOpportunities: opps.records.map(o => ({
        id: o.Id,
        name: o.Name,
        amount: o.Amount,
        stage: o.StageName,
        closeDate: o.CloseDate,
        probability: o.Probability
      }))
    };

  } catch (err) {
    return { source: "Salesforce", error: err.message };
  }
}


// ----------------------
// SERVICENOW CLIENT
// ----------------------
async function fetchServiceNowData() {
  try {
    const url = `${SN_URL}/api/dtp/schindler_txi/incident_summary`;

    const res = await fetch(url, {
      method: "GET",
      headers: {
        "Authorization": "Basic " + Buffer.from(`${SN_USER}:${SN_PASS}`).toString("base64"),
        "Accept": "application/json"
      }
    });

    if (!res.ok) {
      throw new Error(`ServiceNow error ${res.status}`);
    }

    const json = await res.json();
    return json;

  } catch (err) {
    return { source: "ServiceNow", error: err.message };
  }
}


// ----------------------
// SHAREPOINT CLIENT
// ----------------------
async function fetchSharePointData(question) {
  try {
    if (!SP_CHAT_URL) {
      return { source: "SharePoint", error: "SP_CHAT_URL env var not configured." };
    }

    const res = await fetch(SP_CHAT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ question })
    });

    if (!res.ok) {
      throw new Error(`SharePoint returned ${res.status}`);
    }

    return await res.json();

  } catch (err) {
    return { source: "SharePoint", error: err.message };
  }
}


// ----------------------
// GEMINI — COMBINE ANSWER
// ----------------------
async function combineWithGemini(question, sources) {
  try {
    const payload = {
      contents: [
        {
          parts: [
            {
              text: `
You are an enterprise AI assistant. Summarize the three-system data into a leadership-level output.

Question: ${question}

Data:
${JSON.stringify(sources, null, 2)}

Return a clear executive summary with traceability to each system.
`
            }
          ]
        }
      ]
    };

    const gemRes = await fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=" +
        GEMINI_API_KEY,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      }
    );

    const gemJson = await gemRes.json();

    const text =
      gemJson?.candidates?.[0]?.content?.parts?.[0]?.text ||
      "Could not generate answer.";

    return text;

  } catch (err) {
    return "Gemini summarization failed: " + err.message;
  }
}


// ----------------------
// MAIN API HANDLER
// ----------------------
export async function POST(req) {
  try {
    const body = await req.json();
    const question = body.question;

    if (!question) {
      return NextResponse.json(
        { error: "Missing 'question' in request body." },
        { status: 400 }
      );
    }

    // Run all 3 systems in parallel
    const [sf, sn, sp] = await Promise.all([
      fetchSalesforceData(),
      fetchServiceNowData(),
      fetchSharePointData(question)
    ]);

    const sources = { salesforce: sf, serviceNow: sn, sharePoint: sp };

    // Final combined answer from Gemini
    const combinedAnswer = await combineWithGemini(question, sources);

    return NextResponse.json({
      question,
      combinedAnswer,
      sources,
      generatedAt: new Date().toISOString()
    });

  } catch (err) {
    return NextResponse.json(
      { error: err.message || "Server error" },
      { status: 500 }
    );
  }
}
