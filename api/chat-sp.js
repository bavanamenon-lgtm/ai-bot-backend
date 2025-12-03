// api/chat-sp.js
//
// SharePoint -> Graph Drive Search -> Gemini (SAFE MODE)
//
// - Searches only inside the VationGTM site's "Documents" library
// - Summarises only: .docx, .xlsx, .txt, .csv
// - Uses dynamic imports for mammoth/xlsx (no crash if missing)
// - Processes file content in-memory; never stores or logs it.

const {
  GRAPH_TENANT_ID,
  GRAPH_CLIENT_ID,
  GRAPH_CLIENT_SECRET,
  GEMINI_API_KEY,
} = process.env;

// -------- Gemini helper --------

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
- Focus ONLY on this document.
- If the text is very short or looks empty, say that clearly.
- Give a concise summary (5–8 bullet points or short paragraphs).
- Call out any obvious risks, owners, milestones or deadlines if visible.
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

  const raw = await resp.text();
  let data;
  try {
    data = JSON.parse(raw);
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

// -------- Graph token helper --------

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

// -------- Site + Drive helpers (VationGTM) --------

// Hard-coded for this POC:
// Host: vationbangalore.sharepoint.com
// Site path: /sites/VationGTM
const SP_HOSTNAME = "vationbangalore.sharepoint.com";
const SP_SITE_PATH = "/sites/VationGTM";

// Find the site ID and the "Documents" library drive ID for VationGTM
async function getSiteAndDriveIds(accessToken) {
  // 1) Get site
  const siteUrl = `https://graph.microsoft.com/v1.0/sites/${SP_HOSTNAME}:${SP_SITE_PATH}?$select=id`;
  const siteResp = await fetch(siteUrl, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  const siteData = await siteResp.json();
  if (!siteResp.ok) {
    console.error("Graph get site error (masked):", siteData.error || "unknown");
    throw new Error("Failed to get VationGTM SharePoint site from Graph");
  }
  const siteId = siteData.id;
  if (!siteId) {
    throw new Error("VationGTM siteId not found in Graph response");
  }

  // 2) Get drives for the site, pick the "Documents" library
  const drivesUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`;
  const drivesResp = await fetch(drivesUrl, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  const drivesData = await drivesResp.json();
  if (!drivesResp.ok) {
    console.error(
      "Graph get drives error (masked):",
      drivesData.error || "unknown"
    );
    throw new Error("Failed to get drives for VationGTM site");
  }

  if (!Array.isArray(drivesData.value) || !drivesData.value.length) {
    throw new Error("No drives returned for VationGTM site");
  }

  // Prefer the drive named "Documents" or driveType = documentLibrary
  let docsDrive =
    drivesData.value.find((d) => d.name === "Documents") ||
    drivesData.value.find((d) => d.driveType === "documentLibrary") ||
    drivesData.value[0];

  return {
    siteId,
    driveId: docsDrive.id,
  };
}

// -------- Search helpers --------

function getExtension(name = "") {
  const parts = name.split(".");
  if (parts.length < 2) return "";
  return parts[parts.length - 1].toLowerCase();
}

function isSupportedExtension(ext) {
  return ["docx", "xlsx", "txt", "csv"].includes(ext);
}

// Try to pull a "file-like" term from the question
function buildSearchTerm(question) {
  if (!question) return "";

  // 1) Text inside quotes
  const quoted = question.match(/["']([^"']+)["']/);
  if (quoted && quoted[1]) return quoted[1];

  // 2) Phrase before "document" or "file"
  const docMatch = question.match(/summaris\w*\s+(.+?)\s+(document|file)/i);
  if (docMatch && docMatch[1]) return docMatch[1];

  // 3) Fallback to whole question
  return question;
}

// Use Drive search limited to the VationGTM Documents drive
async function searchFileInDrive(question, accessToken) {
  const { driveId } = await getSiteAndDriveIds(accessToken);

  const term = buildSearchTerm(question).slice(0, 200) || "";
  const encoded = encodeURIComponent(term);

  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root/search(q='${encoded}')`;

  const resp = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  const raw = await resp.text();
  let data;
  try {
    data = JSON.parse(raw);
  } catch (e) {
    console.error("Drive search non-JSON response (masked).");
    throw new Error("Drive search returned non-JSON response");
  }

  if (!resp.ok) {
    console.error("Drive search error (masked):", data.error || "unknown");
    throw new Error("Drive search API error");
  }

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
      });
    }
  }

  return results;
}

// -------- Download + extract --------

async function downloadFileBuffer(accessToken, driveId, itemId) {
  if (!driveId || !itemId) {
    throw new Error("Missing driveId or itemId for file download");
  }

  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;

  const resp = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!resp.ok) {
    console.error("Graph file download error (masked):", resp.status);
    throw new Error("Failed to download file content from Graph");
  }

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

  return "";
}

// -------- Main handler --------

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

    // 1) Search inside VationGTM Documents drive
    const results = await searchFileInDrive(question, token);

    if (!results.length) {
      res.status(200).json({
        answer:
          "I couldn't find any matching SharePoint files in the VationGTM Documents library for that question. Try including the exact file name or a key phrase from the document.",
        source: "sharepoint",
        site: SP_SITE_PATH,
        sharePointResults: [],
      });
      return;
    }

    // 2) Pick first supported file
    const supported = results.find((r) =>
      isSupportedExtension(getExtension(r.name))
    );

    if (!supported) {
      res.status(200).json({
        answer:
          "I found files for your search, but none are in a supported format for safe summarisation (allowed: .docx, .xlsx, .txt, .csv).",
        source: "sharepoint",
        site: SP_SITE_PATH,
        sharePointResults: results,
      });
      return;
    }

    const ext = getExtension(supported.name);

    // 3) Download + extract
    const buffer = await downloadFileBuffer(
      token,
      supported.driveId,
      supported.id
    );
    const extractedText = await extractTextFromBuffer(buffer, ext);

    if (!extractedText || !extractedText.trim()) {
      res.status(200).json({
        answer:
          "I could access the file but couldn't safely extract meaningful text from it in this environment. For this POC, only simple text-based documents (txt/csv and some docx/xlsx) are summarised.",
        source: "sharepoint",
        site: SP_SITE_PATH,
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

    // 4) Summarise with Gemini
    const summary = await callGeminiSummary({
      question,
      fileName: supported.name,
      extractedText,
    });

    res.status(200).json({
      answer: summary,
      source: "sharepoint",
      site: SP_SITE_PATH,
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
