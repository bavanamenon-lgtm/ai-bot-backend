// api/chat-sp.js
//
// SharePoint -> Graph -> Gemini (SAFE MODE, crash-proof)
// - Summarises: .docx, .xlsx, .txt, .csv
// - Uses dynamic imports for mammoth/xlsx so missing libs don't crash the function
// - No file content is stored or logged; processed in-memory only.

const {
  GRAPH_TENANT_ID,
  GRAPH_CLIENT_ID,
  GRAPH_CLIENT_SECRET,
  GEMINI_API_KEY,
} = process.env;

// ---------------- Gemini helper ----------------

async function callGeminiSummary({ question, fileName, extractedText }) {
  const MAX_CHARS = 8000;
  const safeText =
    (extractedText || "").toString().slice(0, MAX_CHARS) ||
    "NO_CONTENT_EXTRACTED";

  const url =
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;

  const prompt = `
You are an enterprise-safe assistant summarising a single SharePoint document.

User question:
${question}

File name:
${fileName}

Extracted text from the document (possibly truncated):
${safeText}

Rules:
- Focus only on this document.
- If the text is very short or looks empty, say that clearly.
- Give a concise summary (5–8 bullet points or short paragraphs).
- Call out any obvious risks, deadlines, or owners if visible.
- Do NOT invent data, contacts, or numbers.

Provide the summary now.
`.trim();

  const body = {
    contents: [
      {
        parts: [{ text: prompt }],
      },
    ],
  };

  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  const text = await resp.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    console.error("Gemini non-JSON response (masked).");
    throw new Error("Gemini returned a non-JSON response");
  }

  if (!resp.ok) {
    console.error("Gemini error (masked):", {
      error: data.error || "unknown",
    });
    throw new Error("Gemini API error");
  }

  const candidate = data.candidates && data.candidates[0];
  const parts = candidate && candidate.content && candidate.content.parts;
  const answer = parts && parts[0] && parts[0].text;

  return answer || "I was not able to generate a proper summary.";
}

// ---------------- Graph token helper ----------------

async function getGraphToken() {
  const tenantId = GRAPH_TENANT_ID;
  const clientId = GRAPH_CLIENT_ID;
  const clientSecret = GRAPH_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error(
      "GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET are not set"
    );
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const params = new URLSearchParams();
  params.append("client_id", clientId);
  params.append("client_secret", clientSecret);
  params.append("scope", "https://graph.microsoft.com/.default");
  params.append("grant_type", "client_credentials");

  const resp = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params.toString(),
  });

  const data = await resp.json();
  if (!resp.ok) {
    console.error("Graph token error (masked):", {
      error: data.error || "unknown",
    });
    throw new Error("Failed to get Graph access token");
  }

  return data.access_token;
}

// ---------------- Graph search helper ----------------

const SHAREPOINT_REGION = "IND";

async function searchSharePoint(question, accessToken) {
  const queryString = (question || "").slice(0, 200);

  const url = "https://graph.microsoft.com/v1.0/search/query";

  const body = {
    requests: [
      {
        entityTypes: ["driveItem"],
        query: { queryString },
        from: 0,
        size: 5,
        region: SHAREPOINT_REGION,
      },
    ],
  };

  const resp = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  const raw = await resp.text();
  let data;
  try {
    data = JSON.parse(raw);
  } catch (e) {
    console.error("Graph search non-JSON response (masked).");
    throw new Error("Graph search returned non-JSON response");
  }

  if (!resp.ok) {
    console.error("Graph search error (masked):", {
      error: data.error || "unknown",
    });
    throw new Error(
      `Graph search API error: ${data.error?.code || "Unknown"} - ${
        data.error?.message || "No message"
      }`
    );
  }

  const results = [];
  const hitsContainers =
    data.value &&
    data.value[0] &&
    data.value[0].hitsContainers &&
    data.value[0].hitsContainers[0] &&
    data.value[0].hitsContainers[0].hits;

  if (Array.isArray(hitsContainers)) {
    for (const hit of hitsContainers) {
      const res = hit.resource || {};
      const parent = res.parentReference || {};
      results.push({
        id: res.id || "",
        driveId: parent.driveId || "",
        name: res.name || res.title || "",
        webUrl: res.webUrl || "",
        lastModified: res.lastModifiedDateTime || "",
      });
    }
  }

  return results;
}

// ---------------- Download + extract helpers ----------------

function getExtension(name = "") {
  const parts = name.split(".");
  if (parts.length < 2) return "";
  return parts[parts.length - 1].toLowerCase();
}

function isSupportedExtension(ext) {
  return ["docx", "xlsx", "txt", "csv"].includes(ext);
}

async function downloadFileBuffer(accessToken, driveId, itemId) {
  if (!driveId || !itemId) {
    throw new Error("Missing driveId or itemId for file download");
  }

  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;

  const resp = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (!resp.ok) {
    console.error("Graph file download error (masked):", resp.status);
    throw new Error("Failed to download file content from Graph");
  }

  const arrayBuffer = await resp.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

// Dynamic, crash-proof extraction
async function extractTextFromBuffer(buffer, ext) {
  if (ext === "txt" || ext === "csv") {
    return buffer.toString("utf8");
  }

  if (ext === "docx") {
    try {
      const mammoth = await import("mammoth");
      const result = await mammoth.extractRawText({ buffer });
      return result.value || "";
    } catch (e) {
      console.error("DOCX extraction not available (masked):", e.message);
      return "";
    }
  }

  if (ext === "xlsx") {
    try {
      const xlsxModule = await import("xlsx");
      const XLSX = xlsxModule.default || xlsxModule;
      const workbook = XLSX.read(buffer, { type: "buffer" });
      let textChunks = [];

      workbook.SheetNames.forEach((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) return;
        const sheetJson = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          blankrows: false,
        });
        sheetJson.forEach((row) => {
          const cells = (row || []).map((c) => (c == null ? "" : String(c)));
          const line = cells.join(" | ").trim();
          if (line) textChunks.push(line);
        });
      });

      return textChunks.join("\n");
    } catch (e) {
      console.error("XLSX extraction not available (masked):", e.message);
      return "";
    }
  }

  // Unsupported (shouldn't be called for other types in SAFE mode)
  return "";
}

// ---------------- Main handler ----------------

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res
      .status(405)
      .json({ error: 'Use POST with JSON body { "question": "..." }' });
    return;
  }

  try {
    const { question } = req.body || {};
    if (!question || typeof question !== "string") {
      res.status(400).json({ error: 'Missing "question" in request body.' });
      return;
    }

    const token = await getGraphToken();

    const results = await searchSharePoint(question, token);
    if (!results.length) {
      res.status(200).json({
        answer:
          "I couldn't find any matching SharePoint files for that question. Try a more specific file name or keyword.",
        source: "sharepoint",
        regionUsed: SHAREPOINT_REGION,
        sharePointResults: [],
      });
      return;
    }

    const supported = results.find((r) =>
      isSupportedExtension(getExtension(r.name))
    );

    if (!supported) {
      res.status(200).json({
        answer:
          "I found files for your search, but none of them are in a supported format for safe summarisation (allowed: .docx, .xlsx, .txt, .csv).",
        source: "sharepoint",
        regionUsed: SHAREPOINT_REGION,
        sharePointResults: results,
      });
      return;
    }

    const ext = getExtension(supported.name);

    const buffer = await downloadFileBuffer(token, supported.driveId, supported.id);
    const extractedText = await extractTextFromBuffer(buffer, ext);

    if (!extractedText || !extractedText.trim()) {
      // Very important: graceful, no crash path
      res.status(200).json({
        answer:
          "I could access the file but couldn't safely extract meaningful text from it in this environment. For this POC, only simple text-based documents (txt/csv and some docx/xlsx) are summarised.",
        source: "sharepoint",
        regionUsed: SHAREPOINT_REGION,
        chosenFile: {
          name: supported.name,
          webUrl: supported.webUrl,
          lastModified: supported.lastModified,
          extension: ext,
        },
        sharePointResults: results,
      });
      return;
    }

    const summary = await callGeminiSummary({
      question,
      fileName: supported.name,
      extractedText,
    });

    res.status(200).json({
      answer: summary,
      source: "sharepoint",
      regionUsed: SHAREPOINT_REGION,
      chosenFile: {
        name: supported.name,
        webUrl: supported.webUrl,
        lastModified: supported.lastModified,
        extension: ext,
      },
      sharePointResults: results,
    });
  } catch (err) {
    console.error("Error in /api/chat-sp (masked):", String(err));
    res.status(500).json({
      error: "Internal error in /api/chat-sp.",
      details: String(err),
    });
  }
}
