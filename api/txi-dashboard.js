// /api/txi-dashboard.js

import { GoogleGenerativeAI } from "@google/generative-ai";

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

// Helper: choose model from env, default to 1.5-flash
function getTxiModel() {
  const modelName =
    process.env.GEMINI_TXI_MODEL ||
    "gemini-1.5-flash"; // keep it boring & stable
  return genAI.getGenerativeModel({ model: modelName });
}

// ---- 1. Your existing data fetchers (stubbed here) ----

async function fetchFromServiceNow() {
  // TODO: replace with your real call
  // Return exactly what you’re already returning to the UI
  // I’m using your sample JSON structure:
  return {
    source: "ServiceNow",
    generatedAt: "2025/06/11 11:06:51",
    totalHighPriority: 78,
    byPriority: [
      { priority: "1", count: 72 },
      { priority: "2", count: 6 },
      { priority: "3", count: 105 },
    ],
    ebcIncidents: [],
  };
}

async function fetchFromSalesforce() {
  // TODO: plug in your jsforce/REST logic
  return {
    source: "Salesforce",
    ebcAccount: {
      id: "001gL00000YOfJsQAL",
      name: "EBC HQ",
      industry: "Manufacturing",
      rating: "Hot",
    },
    atRiskSummary: {
      opportunityCount: 2,
      totalAmount: 360000,
    },
    atRiskOpportunities: [
      {
        id: "006gL00000FR428QAD",
        name: "Lift Modernization Deal",
        amount: 240000,
        stage: "Prospecting",
        closeDate: "2025-12-19",
        probability: 10,
      },
      {
        id: "006gL00000FR4i1QAD",
        name: "IoT Annual Renewal",
        amount: 120000,
        stage: "Proposal/Price Quote",
        closeDate: "2025-12-25",
        probability: 30,
      },
    ],
  };
}

async function fetchFromSharePoint(question) {
  // TODO: call your /api/chat-sp endpoint
  // For safety, keep the shape like your current response
  return {
    source: "SharePoint",
    answer:
      "I couldn't find any matching SharePoint files in the VationGTM Documents library for that question. Try using the exact file name or a strong keyword from inside the document.",
    candidateFiles: [],
  };
}

// ---- 2. Fallback summary generator (no Gemini) ----

function buildFallbackSummary({ serviceNow, salesforce, sharePoint, question }) {
  // Use the JSON to build a clean narrative similar to section 1.
  const sn = serviceNow || {};
  const sf = salesforce || {};
  const sp = sharePoint || {};

  const p1 = sn.byPriority?.find((p) => p.priority === "1")?.count ?? 0;
  const p2 = sn.byPriority?.find((p) => p.priority === "2")?.count ?? 0;
  const p3 = sn.byPriority?.find((p) => p.priority === "3")?.count ?? 0;

  const acct = sf.ebcAccount;
  const atRisk = sf.atRiskSummary;

  const lines = [];

  lines.push(
    `Here is a leadership view of today’s biggest risks and impacts based on live system data.`
  );

  // Salesforce
  if (acct && atRisk) {
    lines.push(
      `\n1. Revenue risk on key customer (Source: Salesforce)\n` +
        `   - Account: ${acct.name} (Industry: ${acct.industry}, Rating: ${acct.rating})\n` +
        `   - Open at-risk opportunities: ${atRisk.opportunityCount} deal(s) with total value around $${atRisk.totalAmount}\n`
    );
  }

  // ServiceNow
  if (sn.totalHighPriority != null) {
    lines.push(
      `2. High-priority IT incident load (Source: ServiceNow)\n` +
        `   - High-priority incidents open: ${sn.totalHighPriority}\n` +
        `   - Distribution by priority: P1: ${p1}, P2: ${p2}, P3: ${p3}\n`
    );
  }

  // SharePoint
  if (sp.source) {
    lines.push(
      `3. Knowledge / collaboration signals (Source: SharePoint)\n` +
        `   - Assistant summary: ${sp.answer || "No clear documents found for this question."}`
    );
  }

  return lines.join("");
}

// ---- 3. Main handler ----

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Use POST with JSON body {question: '...'}" });
  }

  try {
    const { question } = req.body || {};
    if (!question || typeof question !== "string") {
      return res
        .status(400)
        .json({ error: "Missing 'question' in request body or not a string." });
    }

    // Fetch all three systems in parallel
    const [serviceNow, salesforce, sharePoint] = await Promise.all([
      fetchFromServiceNow(),
      fetchFromSalesforce(),
      fetchFromSharePoint(question),
    ]);

    let combinedAnswer = null;
    let geminiError = null;

    // 1) Try Gemini
    try {
      const model = getTxiModel();
      const prompt = `
You are an executive analyst. Summarize the organisation's top 3 operational issues
and the business impact, using ONLY the JSON below.

Question from leader:
"${question}"

ServiceNow JSON:
${JSON.stringify(serviceNow, null, 2)}

Salesforce JSON:
${JSON.stringify(salesforce, null, 2)}

SharePoint JSON:
${JSON.stringify(sharePoint, null, 2)}

Return a short leadership-style answer with clear bullets and headings.
`;
      const result = await model.generateContent(prompt);
      combinedAnswer = result.response.text();
    } catch (err) {
      console.error("Gemini error in /api/txi-dashboard:", err);
      geminiError =
        err?.message || "Unknown Gemini error. Falling back to rule-based summary.";
    }

    // 2) If Gemini failed, use fallback
    if (!combinedAnswer) {
      combinedAnswer = buildFallbackSummary({
        serviceNow,
        salesforce,
        sharePoint,
        question,
      });
    }

    return res.status(200).json({
      question,
      combinedAnswer,
      sources: { serviceNow, salesforce, sharePoint },
      geminiError,
      generatedAt: new Date().toISOString(),
    });
  } catch (e) {
    console.error("Fatal error in /api/txi-dashboard:", e);
    return res.status(500).json({ error: "Internal server error in txi-dashboard." });
  }
}
