//
// /api/chat-sp.js
// ENTERPRISE-GRADE SHAREPOINT ASSISTANT (UPDATED WITH SAFE SEARCH TERM SANITISATION)
//
// - Searches VationGTM "Documents" drive via Microsoft Graph
// - Expands folders to find supported files inside
// - Supported types: .txt, .csv, .docx, .xlsx, .pdf
// - Evaluates up to 3 best candidates using Gemini
// - Returns clear summary + which files were considered
//

const {
  GRAPH_TENANT_ID,
  GRAPH_CLIENT_ID,
  GRAPH_CLIENT_SECRET,
  GEMINI_API_KEY,
} = process.env;

// ---------- CORS ----------

function setCorsHeaders(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
}

// ---------- Helpers: extensions, filters ----------

function getExtension(name = "") {
  const parts = name.split(".");
  if (parts.length < 2) return "";
  return parts.pop().toLowerCase();
}

function isSupportedExtension(ext) {
  return ["txt", "csv", "docx", "xlsx", "pdf"].includes(ext);
}

// ---------- Search term extractor (UPDATED, SAFE) ----------

function buildSearchTerm(question) {
  if (!question) return "";
  let q = question.trim();

  // 1) Handle straight + curly quotes → use inside quotes if present
  const quoted = q.match(
    /["'\u201C\u201D\u2018\u2019]([^"'\u201C\u201D\u2018\u2019]+)["'\u201C\u201D\u2018\u2019]/
  );
  if (quoted && quoted[1]) {
    q = quoted[1].trim();
  }

  // 2) Explicit filename pattern
  const filenameMatch = q.match(/([A-Za-z0-9_\- ]+\.[A-Za-z0-9]{1,10})/);
  if (filenameMatch && filenameMatch[1]) {
    q = filenameMatch[1].trim();
  }

  // 3) Phrases like "summarise X document/file"
  const docMatch = q.match(
    /summaris\w*\s+(.+?)\s+(document|file|doc|docs|documents)/i
  );
  if (docMatch && docMatch[1]) {
    q = docMatch[1].trim();
  }

  // 4) Remove common filler prefixes to bias towards the real topic
  const prefixes = [
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
    "do we have any",
    "i need",
    "i want",
  ];
  const lower = q.toLowerCase();
  for (const prefix of prefixes) {
    if (lower.startsWith(prefix + " ")) {
      q = q.slice(prefix.length).trim();
      break;
    }
  }

  // 5) Strip punctuation that can break Graph path or be rejected (? < > # & etc.)
  //    This is what was causing: "A potentially dangerous Request.Path value was detected..."
  q = q.replace(/[?<>#&]/g, " ").replace(/\s+/g, " ").trim();

  // 6) Truncate to avoid insanely long queries
  return q.slice(0, 200);
}

// ---------- Microsoft Graph auth ----------

async function getGraphToken() {
  if (!GRAPH_TENANT_ID || !GRAPH_CLIENT_ID || !GRAPH_CLIENT_SECRET) {
    throw new Error(
      "GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET must be set"
    );
  }

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
    throw new Error("Failed to get Graph access token");
  }

  return data.access_token;
}

// ---------- Site & Drives (VationGTM) ----------

const SP_HOSTNAME = "vationbangalore.sharepoint.com";
const SP_SITE_PATH = "/sites/VationGTM";

async function getSiteAndDriveIds(accessToken) {
  const siteUrl = `https://graph.microsoft.com/v1.0/sites/${SP_HOSTNAME}:${SP_SITE_PATH}?$select=id`;

  const siteResp = await fetch(siteUrl, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  const siteData = await siteResp.json();

  if (!siteResp.ok) {
    console.error("[Graph Site ERROR]:", siteData);
    throw new Error("Failed to resolve VationGTM site in Graph");
  }

  const siteId = siteData.id;
  console.log("[SP] siteId:", siteId);

  const drivesResp = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  const drivesData = await drivesResp.json();

  if (!drivesResp.ok) {
    console.error("[Graph Drives ERROR]:", drivesData);
    throw new Error("Failed to list drives for VationGTM site");
  }

  const drive =
    drivesData.value.find((d) => d.name === "Documents") ||
    drivesData.value.find((d) => d.driveType === "documentLibrary") ||
    drivesData.value[0];

  if (!drive) {
    throw new Error("No drives found for VationGTM site");
  }

  console.log("[SP] driveId:", drive.id);

  return { siteId, driveId: drive.id };
}

// ---------- Graph: search + folder expansion (UPDATED to guard empty term) ----------

async function searchDriveForQuestion(question, token) {
  const { driveId } = await getSiteAndDriveIds(token);

  const term = buildSearchTerm(question);
  if (!term) {
    throw new Error("Search term is empty after sanitisation");
  }

  const encoded = encodeURIComponent(term);

  console.log("[SP] Question:", question);
  console.log("[SP] Search term:", term);

  const url =
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root/search(q='${encoded}')` +
    "?$select=name,id,webUrl,folder,file,lastModifiedDateTime,parentReference";

  const resp = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  const raw = await resp.text();
  let data;
  try {
    data = JSON.parse(raw);
  } catch {
    console.error("[SP] Search non-JSON response:", raw);
    throw new Error("Graph search returned non-JSON response");
  }

  if (!resp.ok) {
    console.error("[SP] Search ERROR:", data);
    throw new Error("Graph search API error");
  }

  const baseResults = Array.isArray(data.value) ? data.value : [];

  console.log(
    "[SP] Raw search hits:",
    baseResults.map((x) => x.name)
  );

  const candidates = [];

  // 1) Files directly returned
  for (const item of baseResults) {
    if (item.file) {
      candidates.push({
        id: item.id,
        driveId: item.parentReference?.driveId || driveId,
        name: item.name,
        webUrl: item.webUrl,
        lastModified: item.lastModifiedDateTime,
        kind: "file",
      });
    }
  }

  // 2) Folders: expand each and collect files inside (up to a limit)
  const folderItems = baseResults.filter((i) => i.folder);
  for (const folder of folderItems) {
    const folderDriveId = folder.parentReference?.driveId || driveId;
    const folderId = folder.id;
    const folderName = folder.name;

    const childrenUrl = `https://graph.microsoft.com/v1.0/drives/${folderDriveId}/items/${folderId}/children?$top=20&$select=name,id,webUrl,file,lastModifiedDateTime,parentReference`;
    const childResp = await fetch(childrenUrl, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const childRaw = await childResp.text();
    let childData;
    try {
      childData = JSON.parse(childRaw);
    } catch {
      console.error(
        `[SP] Folder children non-JSON for ${folderName}:`,
        childRaw
      );
      continue;
    }

    if (!childResp.ok) {
      console.error(`[SP] Folder children ERROR for ${folderName}:`, childData);
      continue;
    }

    const children = Array.isArray(childData.value) ? childData.value : [];

    console.log(
      `[SP] Folder ${folderName} children:`,
      children.map((x) => x.name)
    );

    for (const child of children) {
      if (child.file) {
        candidates.push({
          id: child.id,
          driveId: child.parentReference?.driveId || folderDriveId,
          name: child.name,
          webUrl: child.webUrl,
          lastModified: child.lastModifiedDateTime,
          kind: "file-in-folder",
          parentFolder: folderName,
        });
      }
    }
  }

  return { driveId, candidates };
}

// ---------- Download & extract text ----------

async function downloadFileBuffer(token, driveId, itemId) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;
  const resp = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!resp.ok) {
    console.error("[SP] File download ERROR:", resp.status, resp.statusText);
    throw new Error("Failed to download file content");
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
      const mammothModule = await import("mammoth");
      const mammoth = mammothModule.default || mammothModule;
      const result = await mammoth.extractRawText({ buffer });
      return result.value || "";
    } catch (e) {
      console.error("[DOCX extraction ERROR]:", e);
      return "";
    }
  }

  if (ext === "xlsx") {
    try {
      const xlsxModule = await import("xlsx");
      const XLSX = xlsxModule.default || xlsxModule;
      const workbook = XLSX.read(buffer, { type: "buffer" });
      const lines = [];

      workbook.SheetNames.forEach((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) return;
        const rows = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          blankrows: false,
        });
        rows.forEach((row) => {
          const cells = (row || []).map((c) => (c == null ? "" : String(c)));
          const line = cells.join(" | ").trim();
          if (line) lines.push(line);
        });
      });

      return lines.join("\n");
    } catch (e) {
      console.error("[XLSX extraction ERROR]:", e);
      return "";
    }
  }

  if (ext === "pdf") {
    try {
      const pdfModule = await import("pdf-parse");
      const pdfParse = pdfModule.default || pdfModule;
      const data = await pdfParse(buffer);
      return data.text || "";
    } catch (e) {
      console.error("[PDF extraction ERROR]:", e);
      return "";
    }
  }

  // Unsupported extension
  return "";
}

// ---------- Gemini summariser over multiple docs ----------

async function callGeminiForBestDoc(question, docs) {
  if (!GEMINI_API_KEY) {
    throw new Error("GEMINI_API_KEY must be set");
  }

  const MAX_DOCS = 3;
  const MAX_CHARS_PER_DOC = 6000;

  const limitedDocs = docs.slice(0, MAX_DOCS).map((d, idx) => ({
    index: idx + 1,
    name: d.name,
    text: (d.text || "").slice(0, MAX_CHARS_PER_DOC),
  }));

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" +
    GEMINI_API_KEY;

  const docBlocks = limitedDocs
    .map(
      (d) =>
        `Document ${d.index} — ${d.name}\n` +
        `-----------------------------\n` +
        `${d.text}\n`
    )
    .join("\n\n");

  const prompt = `
You are an enterprise assistant working with multiple SharePoint documents.

User question:
${question}

You have the following candidate documents (possibly truncated):

${docBlocks}

Instructions:
- First, decide which single document is MOST relevant to the user's question.
- If none are relevant, say clearly that you couldn't find a suitable document and briefly mention what you did see.
- If one is relevant, clearly state which document you are using (by name).
- Then provide a concise, accurate summary (5–8 bullet points or short paragraphs).
- Do NOT invent data. If something is not present in the text, do not assume it.
  `.trim();

  const body = {
    contents: [{ parts: [{ text: prompt }] }],
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
    console.error("[Gemini non-JSON]:", raw);
    throw new Error("Gemini returned non-JSON response");
  }

  if (!resp.ok) {
    console.error("[Gemini ERROR]:", data);
    throw new Error("Gemini API error");
  }

  const answer =
    data?.candidates?.[0]?.content?.parts?.[0]?.text ||
    "I was not able to generate a proper summary.";
  return answer;
}

// ---------- Main handler ----------

export default async function handler(req, res) {
  setCorsHeaders(res);

  if (req.method === "OPTIONS") {
    res.status(200).end();
    return;
  }

  if (req.method !== "POST") {
    res
      .status(405)
      .json({ error: 'Use POST with JSON body { "question": "..." }' });
    return;
  }

  try {
    const { question } = req.body || {};
    if (!question || typeof question !== "string") {
      res
        .status(400)
        .json({ error: 'Missing "question" in request body or not a string.' });
      return;
    }

    const token = await getGraphToken();

    const { candidates } = await searchDriveForQuestion(question, token);

    if (!candidates.length) {
      res.status(200).json({
        answer:
          "I couldn't find any matching SharePoint files in the VationGTM Documents library for that question. Try using the exact file name or a strong keyword from inside the document.",
        candidateFiles: [],
      });
      return;
    }

    // Filter to supported extensions
    const fileCandidates = candidates.filter((c) =>
      isSupportedExtension(getExtension(c.name))
    );

    if (!fileCandidates.length) {
      res.status(200).json({
        answer:
          "I found items for your search, but none are in a supported format. Allowed formats are: .txt, .csv, .docx, .xlsx, .pdf.",
        candidateFiles: candidates,
      });
      return;
    }

    // Download & extract text for top few candidates
    const docsWithText = [];
    for (const c of fileCandidates) {
      const ext = getExtension(c.name);
      try {
        const buf = await downloadFileBuffer(token, c.driveId, c.id);
        const text = await extractTextFromBuffer(buf, ext);
        if (text && text.trim()) {
          docsWithText.push({
            ...c,
            extension: ext,
            text,
          });
        } else {
          console.warn(
            `[SP] No extractable text from ${c.name} (${ext}) — skipping`
          );
        }
      } catch (e) {
        console.error(`[SP] Error handling file ${c.name}:`, e);
      }
    }

    if (!docsWithText.length) {
      res.status(200).json({
        answer:
          "I located some files, but I was not able to extract readable text from any of them. They might be empty, image-only, or protected documents.",
        candidateFiles: fileCandidates,
      });
      return;
    }

    // Let Gemini choose best doc and summarise
    const summary = await callGeminiForBestDoc(question, docsWithText);

    res.status(200).json({
      answer: summary,
      usedFiles: docsWithText.map((d) => ({
        name: d.name,
        webUrl: d.webUrl,
        extension: d.extension,
        lastModified: d.lastModified,
        parentFolder: d.parentFolder || null,
      })),
      candidateFiles: candidates,
    });
  } catch (err) {
    console.error("[/api/chat-sp ERROR]:", err);
    res.status(500).json({
      error:
        "Internal server error in SharePoint assistant. Please check logs or configuration.",
    });
  }
}
