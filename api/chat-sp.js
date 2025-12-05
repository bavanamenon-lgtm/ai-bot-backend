//
// FULL REWRITE — SharePoint → Graph → Gemini Summariser
// Drop this entire file into /api/chat-sp.js
//

const {
  GRAPH_TENANT_ID,
  GRAPH_CLIENT_ID,
  GRAPH_CLIENT_SECRET,
  GEMINI_API_KEY,
} = process.env;

/* ----------------------------------------------------
   CORS
---------------------------------------------------- */
function setCorsHeaders(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
}

/* ----------------------------------------------------
   Gemini Summariser
---------------------------------------------------- */
async function callGeminiSummary({ question, fileName, extractedText }) {
  const MAX_CHARS = 8000;
  const safeText = (extractedText || "").slice(0, MAX_CHARS);

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" +
    GEMINI_API_KEY;

  const prompt = `
You are summarising a document retrieved from SharePoint.

User question:
${question}

Document:
${fileName}

Extracted text (may be truncated):
${safeText}

Instructions:
- Summarise accurately
- Do NOT invent data
- Provide 5–8 clear bullet points or short paragraphs
- If text is empty, say so
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
    console.error("[Gemini] NON-JSON RESPONSE:", raw);
    throw new Error("Gemini returned non-JSON");
  }

  if (!resp.ok) {
    console.error("[Gemini] ERROR:", data);
    throw new Error("Gemini API error");
  }

  const parts =
    data?.candidates?.[0]?.content?.parts?.[0]?.text || "No content";
  return parts;
}

/* ----------------------------------------------------
   Microsoft Graph Auth
---------------------------------------------------- */
async function getGraphToken() {
  const url = `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams({
    client_id: GRAPH_CLIENT_ID,
    client_secret: GRAPH_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params.toString(),
  });

  const data = await resp.json();
  if (!resp.ok) {
    console.error("[Graph Token ERROR]:", data);
    throw new Error("Failed to get Graph token");
  }

  return data.access_token;
}

/* ----------------------------------------------------
   SITE + DRIVE Lookup (Your VationGTM site)
---------------------------------------------------- */
const SP_HOSTNAME = "vationbangalore.sharepoint.com";
const SP_SITE_PATH = "/sites/VationGTM";

async function getSiteAndDriveIds(token) {
  const siteUrl = `https://graph.microsoft.com/v1.0/sites/${SP_HOSTNAME}:${SP_SITE_PATH}?$select=id`;

  const siteResp = await fetch(siteUrl, {
    headers: { Authorization: `Bearer ${token}` },
  });

  const siteData = await siteResp.json();
  if (!siteResp.ok) {
    console.error("[Graph Site ERROR]:", siteData);
    throw new Error("Failed to get site ID");
  }

  const siteId = siteData.id;
  console.log("[SP] SiteId:", siteId);

  // Get Drives
  const driveResp = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const driveData = await driveResp.json();

  if (!driveResp.ok) {
    console.error("[Graph Drives ERROR]:", driveData);
    throw new Error("Failed to get drives");
  }

  const drive =
    driveData.value.find((d) => d.name === "Documents") ||
    driveData.value.find((d) => d.driveType === "documentLibrary") ||
    driveData.value[0];

  console.log("[SP] DriveId:", drive.id);

  return { siteId, driveId: drive.id };
}

/* ----------------------------------------------------
   EXTENSION HELPERS
---------------------------------------------------- */
function getExt(name = "") {
  return name.split(".").pop().toLowerCase();
}

function isAllowed(ext) {
  return ["txt", "csv", "docx", "xlsx"].includes(ext);
}

/* ----------------------------------------------------
   SEARCH TERM Extractor (FULL IMPROVED VERSION)
---------------------------------------------------- */
function buildSearchTerm(question) {
  if (!question) return "";
  let q = question.trim();

  // Handle straight + curly quotes
  const quoted = q.match(
    /["'\u201C\u201D\u2018\u2019]([^"'\u201C\u201D\u2018\u2019]+)["'\u201C\u201D\u2018\u2019]/
  );
  if (quoted && quoted[1]) return quoted[1].trim();

  // Filename detection
  const filename = q.match(/([A-Za-z0-9_\-]+\.[A-Za-z0-9]{1,10})/);
  if (filename) return filename[1];

  // Patterns like: summarise X document
  const docMatch = q.match(
    /summaris\w*\s+(.+?)\s+(document|file|doc|docs|documents)/i
  );
  if (docMatch && docMatch[1]) return docMatch[1];

  // Remove leading filler phrases
  const fillers = [
    "can you",
    "could you",
    "please",
    "would you",
    "search",
    "find",
    "show me",
    "give me",
    "share",
    "summarise",
    "summarize",
  ];
  const lower = q.toLowerCase();
  for (const f of fillers) {
    if (lower.startsWith(f + " ")) {
      return q.slice(f.length).trim().slice(0, 200);
    }
  }

  return q.slice(0, 200);
}

/* ----------------------------------------------------
   SEARCH IN DRIVE
---------------------------------------------------- */
async function searchFileInDrive(question, token) {
  const { driveId } = await getSiteAndDriveIds(token);

  const term = buildSearchTerm(question);
  const encoded = encodeURIComponent(term);

  console.log("[SP] Question:", question);
  console.log("[SP] Search term:", term);

  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root/search(q='${encoded}')?$select=name,id,webUrl,lastModifiedDateTime,parentReference`;

  const resp = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  const raw = await resp.text();
  let data;
  try {
    data = JSON.parse(raw);
  } catch {
    console.error("[SP] Non-JSON search response:", raw);
    throw new Error("Search returned non-JSON");
  }

  if (!resp.ok) {
    console.error("[SP] Search ERROR:", data);
    throw new Error("Graph search error");
  }

  const results =
    data.value?.map((item) => ({
      id: item.id,
      driveId: item.parentReference?.driveId || driveId,
      name: item.name,
      webUrl: item.webUrl,
      lastModified: item.lastModifiedDateTime,
    })) || [];

  console.log("[SP] Found items:", results.map((x) => x.name));

  return results;
}

/* ----------------------------------------------------
   DOWNLOAD + EXTRACT
---------------------------------------------------- */
async function downloadFile(token, driveId, itemId) {
  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  if (!resp.ok) throw new Error("Failed downloading file");
  return Buffer.from(await resp.arrayBuffer());
}

async function extractText(buffer, ext) {
  if (ext === "txt" || ext === "csv") return buffer.toString("utf8");

  if (ext === "docx") {
    try {
      const mammoth = await import("mammoth");
      const out = await mammoth.extractRawText({ buffer });
      return out.value || "";
    } catch (e) {
      console.error("[DOCX ERROR]:", e);
      return "";
    }
  }

  if (ext === "xlsx") {
    try {
      const XLSXmod = await import("xlsx");
      const XLSX = XLSXmod.default || XLSXmod;
      const wb = XLSX.read(buffer, { type: "buffer" });
      let lines = [];
      wb.SheetNames.forEach((s) => {
        const sheet = XLSX.utils.sheet_to_json(wb.Sheets[s], {
          header: 1,
        });
        sheet.forEach((row) =>
          lines.push(row.map((c) => (c ? String(c) : "")).join(" | "))
        );
      });
      return lines.join("\n");
    } catch (e) {
      console.error("[XLSX ERROR]:", e);
      return "";
    }
  }

  return "";
}

/* ----------------------------------------------------
   MAIN HANDLER
---------------------------------------------------- */
export default async function handler(req, res) {
  setCorsHeaders(res);

  if (req.method === "OPTIONS") return res.status(200).end();

  if (req.method !== "POST")
    return res.status(405).json({ error: "Use POST" });

  try {
    const { question } = req.body || {};
    if (!question) return res.status(400).json({ error: "Missing question" });

    const token = await getGraphToken();

    const results = await searchFileInDrive(question, token);

    if (!results.length) {
      return res.status(200).json({
        answer:
          "I couldn't find any matching SharePoint files in the VationGTM Documents library. Try using the exact file name or a keyword from inside the document.",
        sharePointResults: [],
      });
    }

    const file = results.find((r) => isAllowed(getExt(r.name)));
    if (!file) {
      return res.status(200).json({
        answer:
          "I found files but none are supported. Allowed formats: .txt, .csv, .docx, .xlsx",
        sharePointResults: results,
      });
    }

    const ext = getExt(file.name);
    const buf = await downloadFile(token, file.driveId, file.id);
    const text = await extractText(buf, ext);

    if (!text.trim()) {
      return res.status(200).json({
        answer:
          "I located the file, but I could not extract readable text from it.",
        chosenFile: file,
      });
    }

    const summary = await callGeminiSummary({
      question,
      fileName: file.name,
      extractedText: text,
    });

    res.status(200).json({
      answer: summary,
      chosenFile: file,
      sharePointResults: results,
    });
  } catch (err) {
    console.error("[/api/chat-sp ERROR]:", err);
    res.status(500).json({ error: "Internal server error" });
  }
}
