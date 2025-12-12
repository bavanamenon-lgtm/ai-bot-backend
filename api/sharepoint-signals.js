// api/sharepoint-signals.js
// Microsoft Graph → SharePoint "Vation GTM" → Documents library → read seeded files
//
// POST { "question": "..." }  (question is optional; used only for future filtering)
//
// ENV:
// MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET

import mammoth from "mammoth";

const JSON_HEADERS = { "Content-Type": "application/json" };

function allowCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
}

async function fetchJson(url, options = {}, timeoutMs = 25000) {
  const controller = new AbortController();
  const id = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const resp = await fetch(url, { ...options, signal: controller.signal });
    const text = await resp.text();
    let json = null;
    try { json = JSON.parse(text); } catch {}
    return { ok: resp.ok, status: resp.status, json, text };
  } finally {
    clearTimeout(id);
  }
}

async function fetchBinary(url, options = {}, timeoutMs = 25000) {
  const controller = new AbortController();
  const id = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const resp = await fetch(url, { ...options, signal: controller.signal });
    const buf = await resp.arrayBuffer();
    return { ok: resp.ok, status: resp.status, buf };
  } finally {
    clearTimeout(id);
  }
}

async function getGraphToken() {
  const tenant = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;

  if (!tenant || !clientId || !clientSecret) {
    throw new Error("Missing MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET");
  }

  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`;

  const body = new URLSearchParams();
  body.set("grant_type", "client_credentials");
  body.set("client_id", clientId);
  body.set("client_secret", clientSecret);
  body.set("scope", "https://graph.microsoft.com/.default");

  const r = await fetchJson(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  }, 25000);

  if (!r.ok) {
    throw new Error(`Token failed HTTP ${r.status}: ${r.text}`);
  }

  const token = r.json?.access_token;
  if (!token) throw new Error("Token missing access_token");
  return token;
}

async function graphGET(token, path) {
  const url = `https://graph.microsoft.com/v1.0${path}`;
  return fetchJson(url, { method: "GET", headers: { Authorization: `Bearer ${token}` } }, 25000);
}

async function graphGETBinary(token, path) {
  const url = `https://graph.microsoft.com/v1.0${path}`;
  return fetchBinary(url, { method: "GET", headers: { Authorization: `Bearer ${token}` } }, 25000);
}

// Find site by search (works even if you don't know hostname/path)
async function findSiteId(token, siteName = "Vation GTM") {
  const r = await graphGET(token, `/sites?search=${encodeURIComponent(siteName)}`);
  if (!r.ok) throw new Error(`Site search failed HTTP ${r.status}: ${r.text}`);
  const sites = r.json?.value || [];
  const exact = sites.find(s => String(s.name || "").toLowerCase() === siteName.toLowerCase());
  const best = exact || sites[0];
  if (!best?.id) throw new Error(`No site found for search "${siteName}"`);
  return best.id;
}

async function findDriveId(token, siteId, driveName = "Documents") {
  const r = await graphGET(token, `/sites/${encodeURIComponent(siteId)}/drives`);
  if (!r.ok) throw new Error(`Drives fetch failed HTTP ${r.status}: ${r.text}`);
  const drives = r.json?.value || [];
  const match = drives.find(d => String(d.name || "").toLowerCase() === driveName.toLowerCase());
  const best = match || drives[0];
  if (!best?.id) throw new Error(`No drive found on site for "${driveName}"`);
  return best.id;
}

async function tryGetDriveItemByPath(token, driveId, path) {
  // /drives/{drive-id}/root:/path:/content
  const safePath = path.split("/").map(encodeURIComponent).join("/").replace(/%2F/g, "/");
  const meta = await graphGET(token, `/drives/${encodeURIComponent(driveId)}/root:/${safePath}`);
  if (!meta.ok) return { ok: false, status: meta.status, meta: null };

  const item = meta.json;
  if (!item?.id) return { ok: false, status: meta.status, meta: null };
  return { ok: true, status: 200, meta: item };
}

async function readFileContent(token, driveId, itemMeta) {
  const name = itemMeta.name || "";
  const mime = itemMeta.file?.mimeType || "";

  // Use /content for binary/text download
  const bin = await graphGETBinary(token, `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemMeta.id)}/content`);
  if (!bin.ok) throw new Error(`Content download failed for ${name} HTTP ${bin.status}`);

  const buffer = Buffer.from(bin.buf);

  // TXT
  if (name.toLowerCase().endsWith(".txt") || mime.startsWith("text/")) {
    return buffer.toString("utf-8");
  }

  // DOCX
  if (name.toLowerCase().endsWith(".docx") || mime === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
    const out = await mammoth.extractRawText({ buffer });
    return String(out.value || "");
  }

  // Fallback: return empty (we’re not doing PDF parsing here to keep it stable)
  return "";
}

function clip(text, max = 8000) {
  const t = String(text || "").replace(/\r\n/g, "\n").trim();
  if (t.length <= max) return t;
  return t.slice(0, max) + "\n…(truncated)";
}

export default async function handler(req, res) {
  allowCors(res);

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "POST only" });

  try {
    const token = await getGraphToken();
    const siteId = await findSiteId(token, "Vation GTM");
    const driveId = await findDriveId(token, siteId, "Documents");

    const targets = [
      "Annual EBC Review Notes.txt",
      "EBC_Account_Health_Risk.docx",
      "IT_Operations_Weekly_Report.docx",
      "Sales_Risk_Accounts_List.docx"
    ];

    // We check both root and /General/
    const candidatePaths = [];
    for (const f of targets) {
      candidatePaths.push(f);
      candidatePaths.push(`General/${f}`);
    }

    const filesFound = [];
    let combinedText = "";

    for (const p of candidatePaths) {
      const hit = await tryGetDriveItemByPath(token, driveId, p);
      if (!hit.ok) continue;

      const meta = hit.meta;
      const already = filesFound.find(x => x.id === meta.id);
      if (already) continue;

      const content = await readFileContent(token, driveId, meta);
      filesFound.push({ name: meta.name, pathTried: p, id: meta.id, size: meta.size || null });

      // Only append if we extracted usable text
      if (content && content.trim()) {
        combinedText += `\n\n===== ${meta.name} =====\n` + content.trim();
      }
    }

    if (!filesFound.length) {
      return res.status(200).json({
        source: "SharePoint",
        ok: false,
        error: "NO_MATCH: Could not locate seeded files in Vation GTM → Documents (root or General). Check permissions and drive mapping.",
        filesFound: [],
        signalsText: ""
      });
    }

    // If we found files but couldn’t extract text (e.g., PDFs), still return filesFound so you can prove access
    return res.status(200).json({
      source: "SharePoint",
      ok: combinedText.trim().length > 0,
      error: combinedText.trim().length > 0 ? null : "FOUND_FILES_BUT_NO_TEXT: Files exist but text extraction returned empty (check file types/permissions).",
      filesFound,
      signalsText: clip(combinedText, 8000)
    });
  } catch (e) {
    return res.status(200).json({
      source: "SharePoint",
      ok: false,
      error: e?.message || String(e),
      filesFound: [],
      signalsText: ""
    });
  }
}
