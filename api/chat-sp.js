// api/chat-sp.js
//
// Backend for SharePoint -> Graph -> Gemini
// Uses client credentials flow with Microsoft Graph
// and Gemini 2.0 Flash for summarisation.

const {
  GRAPH_TENANT_ID,
  GRAPH_CLIENT_ID,
  GRAPH_CLIENT_SECRET,
  GEMINI_API_KEY,
} = process.env;

// --- Helper: call Gemini ---
async function callGemini(prompt) {
  const url =
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;

  const body = {
    contents: [
      {
        parts: [{ text: prompt }],
      },
    ],
  };

  const resp = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });

  const text = await resp.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    console.error('Gemini non-JSON response:', text);
    throw new Error('Gemini returned a non-JSON response');
  }

  if (!resp.ok) {
    console.error('Gemini error:', data);
    throw new Error('Gemini API error');
  }

  const candidate = data.candidates && data.candidates[0];
  const parts = candidate && candidate.content && candidate.content.parts;
  const answer = parts && parts[0] && parts[0].text;

  return answer || 'I was not able to generate a proper response.';
}

// --- Helper: get Microsoft Graph token (client credentials) ---
async function getGraphToken() {
  const tenantId = GRAPH_TENANT_ID;
  const clientId = GRAPH_CLIENT_ID;
  const clientSecret = GRAPH_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error(
      'GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET are not set'
    );
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const params = new URLSearchParams();
  params.append('client_id', clientId);
  params.append('client_secret', clientSecret);
  params.append('scope', 'https://graph.microsoft.com/.default');
  params.append('grant_type', 'client_credentials');

  const resp = await fetch(tokenUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString(),
  });

  const data = await resp.json();
  if (!resp.ok) {
    console.error('Graph token error:', data);
    throw new Error('Failed to get Graph access token');
  }

  return data.access_token;
}

// --- Helper: search SharePoint content using Graph Search API ---
async function searchSharePoint(queryText, accessToken) {
  // Keep the query short-ish for Graph search
  const queryString = (queryText || '').slice(0, 200);

  const url = 'https://graph.microsoft.com/v1.0/search/query';

  const body = {
    requests: [
      {
        entityTypes: ['driveItem', 'listItem', 'sitePage'],
        query: { queryString },
        from: 0,
        size: 5, // top 5 hits is enough for Gemini
      },
    ],
  };

  const resp = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });

  const data = await resp.json();

  if (!resp.ok) {
    console.error('Graph search error:', data);
    throw new Error('Graph search API error');
  }

  const results = [];

  const containers =
    data.value &&
    data.value[0] &&
    data.value[0].hitsContainers &&
    data.value[0].hitsContainers[0] &&
    data.value[0].hitsContainers[0].hits;

  if (Array.isArray(containers)) {
    for (const hit of containers) {
      const res = hit.resource || {};
      results.push({
        name: res.name || res.title || '',
        webUrl: res.webUrl || '',
        summary: res.description || res.snippet || '',
        lastModified: res.lastModifiedDateTime || '',
        driveType: res.driveType || '',
      });
    }
  }

  return results;
}

// --- Main handler ---
export default async function handler(req, res) {
  if (req.method !== 'POST') {
    res
      .status(405)
      .json({ error: 'Use POST with JSON body { "question": "..." }' });
    return;
  }

  try {
    const { question } = req.body || {};
    if (!question || typeof question !== 'string') {
      res.status(400).json({ error: 'Missing "question" in request body.' });
      return;
    }

    // 1) Get Graph token
    const token = await getGraphToken();

    // 2) Search SharePoint
    const spResults = await searchSharePoint(question, token);

    // 3) Build Gemini prompt
    const prompt = `
You are an assistant helping the user understand information stored in SharePoint.

User question:
${question}

Here are the top SharePoint search results in JSON:
${JSON.stringify(spResults, null, 2)}

Using ONLY the information above:
- Summarise what is relevant to the user's question.
- Mention document names and high-level insights.
- If the results are empty or irrelevant, say so clearly and suggest a better query.

Answer in clear, simple English.
`.trim();

    // 4) Call Gemini for the final answer
    const answer = await callGemini(prompt);

    // 5) Return answer + raw search results
    res.status(200).json({
      answer,
      source: 'sharepoint',
      sharePointResults: spResults,
    });
  } catch (err) {
    console.error('Error in /api/chat-sp:', err);
    res.status(500).json({
      error: 'Internal error in /api/chat-sp.',
      details: String(err),
    });
  }
}
