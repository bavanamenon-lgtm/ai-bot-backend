// txi-dashboard-ui.js
(function () {
  const $ = (id) => document.getElementById(id);

  const qEl = $("q");
  const askBtn = $("askBtn");
  const exampleBtn = $("exampleBtn");
  const toggleDebugBtn = $("toggleDebugBtn");

  const answerEl = $("answer");
  const debugBox = $("debugBox");

  const sfDot = $("sfDot"), snDot = $("snDot"), spDot = $("spDot");
  const sfTxt = $("sfTxt"), snTxt = $("snTxt"), spTxt = $("spTxt");

  function setDot(dotEl, status) {
    dotEl.classList.remove("ok", "err", "warn");
    if (status === "OK") dotEl.classList.add("ok");
    else if (status === "error") dotEl.classList.add("err");
    else dotEl.classList.add("warn");
  }

  function systemStatus(srcObj) {
    if (!srcObj) return "warn";
    if (srcObj.ok === true) return "OK";
    if (srcObj.ok === false) return "error";
    if (srcObj.error) return "error";
    return "warn";
  }

  function toExecFormat(combinedAnswer, sources) {
    let text = (combinedAnswer || "").trim();

    text = text
      .replace(/\r\n/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .replace(/(\n\d\))/g, "\n\n$1");

    if (!/Traceability:/i.test(text)) {
      const sf = sources?.salesforce?.ok ? "Salesforce(OK)" : "Salesforce(error)";
      const sn = sources?.serviceNow?.ok ? "ServiceNow(OK)" : "ServiceNow(error)";
      const sp = sources?.sharePoint?.ok ? "SharePoint(OK)" : "SharePoint(error)";
      text += `\n\nTraceability: ${sf} | ${sn} | ${sp}`;
    }

    return text;
  }

  async function ask() {
    const question = (qEl.value || "").trim();
    if (!question) return;

    askBtn.disabled = true;
    answerEl.classList.add("muted");
    answerEl.textContent = "Working on it…";

    // Ensure answer area never visually clips
    answerEl.style.whiteSpace = "pre-wrap";
    answerEl.style.wordBreak = "break-word";
    answerEl.style.maxHeight = "70vh";
    answerEl.style.overflow = "auto";

    // Reset chips while loading
    setDot(sfDot, "warn"); setDot(snDot, "warn"); setDot(spDot, "warn");
    sfTxt.textContent = "—"; snTxt.textContent = "—"; spTxt.textContent = "—";

    try {
      const r = await fetch("/api/txi-dashboard", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ question })
      });

      const rawText = await r.text();
      let json = null;
      try { json = JSON.parse(rawText); } catch { json = null; }

      if (!r.ok || !json) {
        answerEl.classList.remove("muted");
        answerEl.textContent = `Backend error: ${r.status}\n\n${rawText}`;
        debugBox.textContent = rawText;
        return;
      }

      const sources = json.sources || {};
      const sfS = systemStatus(sources.salesforce);
      const snS = systemStatus(sources.serviceNow);
      const spS = systemStatus(sources.sharePoint);

      setDot(sfDot, sfS); setDot(snDot, snS); setDot(spDot, spS);

      sfTxt.textContent = sfS === "OK" ? "• OK" : sfS === "error" ? "• error" : "• warn";
      snTxt.textContent = snS === "OK" ? "• OK" : snS === "error" ? "• error" : "• warn";
      spTxt.textContent = spS === "OK" ? "• OK" : spS === "error" ? "• error" : "• warn";

      const combined = toExecFormat(json.combinedAnswer, sources);

      answerEl.classList.remove("muted");
      answerEl.textContent = combined;

      // PROOF: not truncating
      console.log("combinedAnswer length:", combined.length);
      console.log("combinedAnswer tail:", combined.slice(-200));

      // Debug (optional)
      const debugPayload = {
        httpStatus: r.status,
        generatedAt: json.generatedAt,
        gemini: json.gemini,
        sources: {
          salesforce: sources.salesforce?.ok ? { ok: true } : { ok: false, error: sources.salesforce?.error || "unknown" },
          serviceNow: sources.serviceNow?.ok ? { ok: true } : { ok: false, error: sources.serviceNow?.error || "unknown" },
          sharePoint: sources.sharePoint?.ok ? { ok: true } : { ok: false, error: sources.sharePoint?.error || "unknown" }
        }
      };
      debugBox.textContent = JSON.stringify(debugPayload, null, 2);

    } catch (e) {
      answerEl.classList.remove("muted");
      answerEl.textContent = `Client error:\n${e?.message || String(e)}`;
      debugBox.textContent = String(e?.stack || e);
    } finally {
      askBtn.disabled = false;
    }
  }

  askBtn.addEventListener("click", ask);

  exampleBtn.addEventListener("click", () => {
    qEl.value = "What are the top 3 operational issues I should care about today, and what’s the business impact?";
  });

  toggleDebugBtn.addEventListener("click", () => {
    const isHidden = (debugBox.style.display === "" || debugBox.style.display === "none");
    debugBox.style.display = isHidden ? "block" : "none";
  });
})();
