// api/chat-sp.js
//
// SharePoint -> Graph Drive Search -> Extract -> Gemini summary (SAFE MODE)
//
// Fixes included:
// - Sanitise search term (avoid dangerous characters like ? ' etc).
// - Deterministic TXI keyword search for leadership questions.
// - Summarise top N files (not only one).
// - Always return JSON.

const {
  GRAPH_TENANT_ID,
  GRAPH_CLIENT_ID,
  GRAPH_CLIENT_SECRET,
  GEMINI_API_KEY,
  GEMINI_MODEL, // optional
} = process.env;

export const config = { runtime: "nodejs" };

// -------- Simple CORS helpers --------
function setCorsHeaders(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
}

// -------- Gemini helper --------
async function callGeminiSummary({ question, files }) {
  const model = GEMINI_MODEL || "gemini-2.0-flash";

  const url =
    `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${GEMINI_API_KEY}`;

  const compactFiles = (files || []).map((f) => ({
    fileName: f.fileName,
    excerpt: (f.extractedText || "").slice(0, 2000),
  }));

  const prompt = `
You are an enterprise-safe assistant summarising SharePoint documents.

User question:
${question}

You have up to ${compactFiles.length} documents. Summarise ONLY what is present.

Rules:
- Output concise bullets.
- If multiple files: group by file name.
- Do NOT invent contacts, numbers, owners.
- If content is empty, say that clearly.
`.trim();

  const body = {
    contents: [
      {
        parts: [{ text: `${prompt}\n\nDOCS:\n${JSON.stringify(compactFiles)}` }],
      },
    ],
  };

  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  const raw = await resp.text();
  let data;
  try {
    data = JSON.parse(raw);
  } catch {
    throw new Error("Gemini returned non-JSON");
  }

  if (!resp.ok) {
    const msg = data?.error?.message || "Gemini API error";
    const code = data?.error?.code || resp.status;
    throw new Error(`Gemini error (${code}): ${msg}`);
  }

  return data?.candidates?.[0]?.content?.parts?.[0]?.text || "";
}

// -------- Graph token helper --------
async function getGraphToken() {
  if (!GRAPH_TENANT_ID || !GRAPH_CLIENT_ID || !GRAPH_CLIENT_SECRET) {
    throw new Error("GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET are not set");
  }

  const tokenUrl = `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append("client_id", GRAPH_CLIENT_ID);
  params.append("client_secret", GRAPH_CLIENT_SECRET);
  params.append("scope", "https://graph.microsoft.com/.default");
  params.append("grant_type", "client_credentials");

  const resp = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params.toString(),
  });

  const data = await resp.json();
  if (!resp.ok) throw new Error("Failed to get Graph access token");
  return data.access_token;
}

// -------- Site + Drive helpers --------
const SP_HOSTNAME = "vationbangalore.sharepoint.com";
const SP_SITE_PATH = "/sites/VationGTM";

async function getSiteAndDriveIds(accessToken) {
  const siteUrl = `https://graph.microsoft.com/v1.0/sites/${SP_HOSTNAME}:${SP_SITE_PATH}?$select=id`;
  const siteResp = await fetch(siteUrl, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  const siteData = await siteResp.json();
  if (!siteResp.ok) throw new Error("Failed to get VationGTM SharePoint site from Graph");
  const siteId = siteData.id;

  const drivesUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`;
  const drivesResp = await fetch(drivesUrl, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  const drivesData = await drivesResp.json();
  if (!drivesResp.ok) throw new Error("Failed to get drives for VationGTM site");

  const docsDrive =
    drivesData.value.find((d) => d.name === "Documents") ||
    drivesData.value.find((d) => d.driveType === "documentLibrary") ||
    drivesData.value[0];

  return { siteId, driveId: docsDrive.id };
}

// -------- Search helpers --------
function getExtension(name = "") {
  const parts = name.split(".");
  return parts.length < 2 ? "" : parts[parts.length - 1].toLowerCase();
}
function isSupportedExtension(ext) {
  return ["docx", "xlsx", "txt", "csv"].includes(ext);
}

// Fix the “dangerous (?)” and other Graph search oddities.
// Graph drive search is sensitive. Keep it alphanumeric-ish.
function sanitizeSearchTerm(term) {
  return (term || "")
    .replace(/[\?\#\%\&\{\}\|\\\^~\[\]`]/g, " ")
    .replace(/['"]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 180);
}

function buildSmartSearchTerm(question) {
  const q = (question || "").toLowerCase();

  // If they ask leadership risk/impact questions, force TXI keywords
  if (q.includes("risk") || q.includes("impact") || q.includes("issues") || q.includes("customer")) {
    return "EBC_Account_Health_Risk IT_Operations_Weekly_Report Sales_Risk_Accounts_List";
  }

  // If they mention a filename explicitly
  const fileMatch = question.match(/([A-Za-z0-9_\- ]+\.(docx|xlsx|txt|csv))/i);
  if (fileMatch && fileMatch[1]) return fileMatch[1];

  return question;
}

async function searchFiles(question, accessToken) {
  const { driveId } = await getSiteAndDriveIds(accessToken);

  const term = sanitizeSearchTerm(buildSmartSearchTerm(question));
  const encoded = encodeURIComponent(term);

  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root/search(q='${encoded}')`;

  const resp = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  const raw = await resp.text();
  let data;
  try {
    data = JSON.parse(raw);
  } catch {
    throw new Error("Drive search returned non-JSON response");
  }

  if (!resp.ok) throw new Error("Drive search API error");

  const results = [];
  if (Array.isArray(data.value)) {
    for (const item of data.value) {
      const parent = item.parentReference || {};
      results.push({
        id: item.id || "",
        driveId: parent.driveId || driveId,
        name: item.name || "",
        webUrl: item.webUrl || "",
        lastModified: item.lastModifiedDateTime || "",
        kind: item.file ? "file" : "other",
      });
    }
  }
  return results;
}

// -------- Download + extract --------
async function downloadFileBuffer(accessToken, driveId, itemId) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;
  const resp = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!resp.ok) throw new Error("Failed to download file content from Graph");
  const arrayBuffer = await resp.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

async function extractTextFromBuffer(buffer, ext) {
  if (ext === "txt" || ext === "csv") {
    return buffer.toString("utf8");
  }

  if (ext === "docx") {
    try {
      const mammoth = await import("mammoth");
      const result = await mammoth.extractRawText({ buffer });
      return result.value || "";
    } catch {
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
        const sheetJson = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
        sheetJson.forEach((row) => {
          const cells = (row || []).map((c) => (c == null ? "" : String(c)));
          const line = cells.join(" | ").trim();
          if (line) textChunks.push(line);
        });
      });

      return textChunks.join("\n");
    } catch {
      return "";
    }
  }

  return "";
}

// -------- Main handler --------
export default async function handler(req, res) {
  setCorsHeaders(res);

  if (req.method === "OPTIONS") {
    res.status(200).end();
    return;
  }

  if (req.method !== "POST") {
    res.status(405).json({ error: 'Use POST with JSON body { "question": "..." }' });
    return;
  }

  try {
    const { question } = req.body || {};
    if (!question || typeof question !== "string") {
      res.status(400).json({ error: 'Missing "question" in request body.' });
      return;
    }

    const token = await getGraphToken();
    const results = await searchFiles(question, token);

    if (!results.length) {
      res.status(200).json({
        answer:
          "I couldn't find any matching SharePoint files in the VationGTM Documents library. Try using an exact file name like 'EBC_Account_Health_Risk.docx' or a strong keyword from inside the document.",
        usedFiles: [],
        candidateFiles: [],
      });
      return;
    }

    // Pick top 3 supported files
    const supported = results
      .filter((r) => isSupportedExtension(getExtension(r.name)))
      .slice(0, 3);

    if (!supported.length) {
      res.status(200).json({
        answer:
          "I found files for your search, but none are in a supported format for summarisation (allowed: .docx, .xlsx, .txt, .csv).",
        usedFiles: [],
        candidateFiles: results,
      });
      return;
    }

    const extracted = [];
    for (const f of supported) {
      const ext = getExtension(f.name);
      const buffer = await downloadFileBuffer(token, f.driveId, f.id);
      const extractedText = await extractTextFromBuffer(buffer, ext);
      extracted.push({
        fileName: f.name,
        extractedText,
        webUrl: f.webUrl,
        lastModified: f.lastModified,
        extension: ext,
      });
    }

    // If Gemini not configured, return deterministic info anyway
    if (!GEMINI_API_KEY) {
      res.status(200).json({
        answer:
          "Files found, but Gemini summarisation is not configured (GEMINI_API_KEY missing).",
        usedFiles: extracted.map((f) => ({
          name: f.fileName,
          webUrl: f.webUrl,
          extension: f.extension,
          lastModified: f.lastModified,
        })),
        candidateFiles: results,
      });
      return;
    }

    const summary = await callGeminiSummary({ question, files: extracted });

    res.status(200).json({
      answer: summary || "I couldn't generate a summary.",
      usedFiles: extracted.map((f) => ({
        name: f.fileName,
        webUrl: f.webUrl,
        extension: f.extension,
        lastModified: f.lastModified,
      })),
      candidateFiles: results,
    });
  } catch (err) {
    res.status(500).json({
      error: "Internal server error in SharePoint assistant. Please check logs or configuration.",
      details: String(err),
    });
  }
}
