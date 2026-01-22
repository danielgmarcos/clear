(function () {
  const defaultApiUrl = "https://cleardrop.wit-software.com/analyze";
  const uiVersion = "0.1.0";
  const analyzeButton = document.getElementById("analyze");
  const apiInput = document.getElementById("api-url");
  const statusEl = document.getElementById("status");
  const resultEl = document.getElementById("result");
  const verdictCard = document.querySelector(".verdict-card");
  const verdictPill = document.getElementById("verdict-pill");
  const verdictTitle = document.getElementById("verdict-title");
  const verdictSubtitle = document.getElementById("verdict-subtitle");
  const versionEl = document.getElementById("ui-version");
  const resultCard = document.getElementById("result-card");
  const toggleResultButton = document.getElementById("toggle-result");

  function setStatus(message, isError) {
    statusEl.textContent = message;
    statusEl.classList.toggle("muted", !isError);
    statusEl.style.color = isError ? "#b00020" : "";
  }

  function setResult(value) {
    resultEl.textContent = value || "";
  }

  function updateVerdict(score, verdict) {
    if (!verdictCard || !verdictPill || !verdictTitle || !verdictSubtitle) {
      return;
    }

    if (typeof score !== "number" || Number.isNaN(score)) {
      verdictCard.dataset.state = "idle";
      verdictPill.textContent = verdict || "Awaiting analysis";
      verdictPill.className = "verdict-pill neutral";
      verdictTitle.textContent = "No scan yet";
      verdictSubtitle.textContent = "Run an analysis to see the verdict.";
      return;
    }

    const clamped = Math.max(0, Math.min(100, score));
    const verdictText =
      verdict || (clamped < 35 ? "Safe" : clamped < 70 ? "Suspicious" : "Phishing");
    verdictPill.textContent = verdictText;
    verdictTitle.textContent = verdictText;
    verdictSubtitle.textContent = `Risk score: ${clamped}%`;

    if (/safe/i.test(verdictText) || clamped < 35) {
      verdictCard.dataset.state = "safe";
      verdictPill.className = "verdict-pill safe";
    } else if (/suspicious|warning/i.test(verdictText) || clamped < 70) {
      verdictCard.dataset.state = "warning";
      verdictPill.className = "verdict-pill warning";
    } else {
      verdictCard.dataset.state = "danger";
      verdictPill.className = "verdict-pill danger";
    }
  }

  function extractScore(payload) {
    if (!payload || typeof payload !== "object") {
      return null;
    }
    const score = Number(payload.model_confidence ?? payload.rules_score);
    return Number.isFinite(score) ? score : null;
  }

  function extractVerdict(payload) {
    if (!payload || typeof payload !== "object") {
      return "";
    }
    return (
      payload.final_verdict ||
      payload.rules_verdict ||
      payload.model_verdict ||
      ""
    );
  }

  function getApiUrl() {
    return (apiInput.value || "").trim();
  }

  function loadApiUrl() {
    apiInput.value = defaultApiUrl;
  }

  function updateAnalyzeState() {
    const hasItem = !!Office.context.mailbox.item;
    const hasApi = !!getApiUrl();
    analyzeButton.disabled = !(hasItem && hasApi);
  }

  function getRecipients(field) {
    if (!field || !Array.isArray(field)) {
      return [];
    }
    return field.map((entry) => ({
      name: entry.displayName || "",
      address: entry.emailAddress || "",
    }));
  }

  function getBodyText() {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Text,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value || "");
          } else {
            reject(result.error || new Error("Failed to read body"));
          }
        }
      );
    });
  }

  function getBodyHtml() {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Html,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value || "");
          } else {
            reject(result.error || new Error("Failed to read body HTML"));
          }
        }
      );
    });
  }

  function extractTextFromHtml(html) {
    if (!html) {
      return "";
    }
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, "text/html");
    return (doc.body && doc.body.textContent ? doc.body.textContent : "").trim();
  }

  function extractLinksFromHtml(html) {
    if (!html) {
      return [];
    }
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, "text/html");
    const anchors = Array.from(doc.querySelectorAll("a[href]"));
    const hrefs = anchors.map((anchor) => anchor.getAttribute("href") || "").filter(Boolean);
    return Array.from(new Set(hrefs));
  }

  function extractLinksFromText(text) {
    if (!text) {
      return [];
    }
    const urlRegex = /\bhttps?:\/\/[^\s<>"')]+/gi;
    return Array.from(new Set(text.match(urlRegex) || []));
  }

  async function collectBody() {
    let bodyHtml = "";
    let bodyText = "";

    try {
      bodyHtml = await getBodyHtml();
    } catch (error) {
      bodyHtml = "";
    }

    try {
      bodyText = await getBodyText();
    } catch (error) {
      bodyText = "";
    }

    if (!bodyText && bodyHtml) {
      bodyText = extractTextFromHtml(bodyHtml);
    }

    const links = bodyHtml
      ? extractLinksFromHtml(bodyHtml)
      : extractLinksFromText(bodyText);

    return { bodyHtml, bodyText, links };
  }

  async function collectAttachments(item) {
    const attachments = Array.isArray(item.attachments) ? item.attachments : [];
    if (!attachments.length) {
      return [];
    }

    if (typeof item.getAttachmentContentAsync !== "function") {
      return attachments.map((att) => ({
        id: att.id,
        name: att.name,
        size: att.size,
        contentType: att.contentType,
        isInline: att.isInline,
        content: null,
        contentFormat: null,
        error: "Attachment content API not available.",
      }));
    }

    const results = await Promise.all(
      attachments.map(
        (att) =>
          new Promise((resolve) => {
            item.getAttachmentContentAsync(att.id, (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve({
                  id: att.id,
                  name: att.name,
                  size: att.size,
                  contentType: att.contentType,
                  isInline: att.isInline,
                  content: result.value.content,
                  contentFormat: result.value.format,
                });
              } else {
                resolve({
                  id: att.id,
                  name: att.name,
                  size: att.size,
                  contentType: att.contentType,
                  isInline: att.isInline,
                  content: null,
                  contentFormat: null,
                  error: result.error ? result.error.message : "Attachment fetch failed.",
                });
              }
            });
          })
      )
    );

    return results;
  }

  async function analyzeEmail() {
    const apiUrl = getApiUrl();
    if (!apiUrl) {
      setStatus("Add the API URL in preferences first.", true);
      return;
    }

    const item = Office.context.mailbox.item;
    if (!item) {
      setStatus("No email is selected.", true);
      return;
    }

    setStatus("Collecting email content...");
    setResult("");
    updateVerdict(null, "Collecting email...");

    let bodyText = "";
    let bodyHtml = "";
    let links = [];
    let attachments = [];
    try {
      const bodyInfo = await collectBody();
      bodyText = bodyInfo.bodyText;
      bodyHtml = bodyInfo.bodyHtml;
      links = bodyInfo.links;
    } catch (error) {
      setStatus("Could not read email body.", true);
      return;
    }

    try {
      attachments = await collectAttachments(item);
    } catch (error) {
      attachments = [];
    }

    const payload = {
      subject: item.subject || "",
      from: item.from
        ? { name: item.from.displayName || "", address: item.from.emailAddress || "" }
        : null,
      to: getRecipients(item.to),
      cc: getRecipients(item.cc),
      bcc: getRecipients(item.bcc),
      itemId: item.itemId || "",
      internetMessageId: item.internetMessageId || "",
      bodyText,
      bodyHtml,
      links,
      attachments,
      receivedDateTime: item.dateTimeCreated || "",
    };

    setStatus("Sending for analysis...");
    updateVerdict(null, "Analyzing...");

    try {
      const response = await fetch(apiUrl, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      const text = await response.text();
      if (!response.ok) {
        setStatus(`API error (${response.status}).`, true);
        setResult(text);
        updateVerdict(null, "Analysis failed");
        return;
      }

      let output = text;
      let parsed = null;
      try {
        parsed = JSON.parse(text);
        output = JSON.stringify(parsed, null, 2);
      } catch (error) {
        // keep raw text
      }

      setStatus("Analysis complete.");
      setResult(output);
      updateVerdict(extractScore(parsed), extractVerdict(parsed));
    } catch (error) {
      setStatus("Network error sending to API.", true);
      setResult(String(error));
      updateVerdict(null, "Network error");
    }
  }

  Office.onReady(() => {
    loadApiUrl();
    updateAnalyzeState();
    updateVerdict(null, "Awaiting analysis");
    if (versionEl) {
      versionEl.textContent = `v${uiVersion}`;
    }
    if (toggleResultButton && resultCard) {
      toggleResultButton.addEventListener("click", () => {
        const isHidden = resultCard.classList.toggle("is-hidden");
        toggleResultButton.textContent = isHidden ? "See reasoning" : "Hide reasoning";
      });
    }
    analyzeButton.addEventListener("click", analyzeEmail);
    apiInput.addEventListener("input", updateAnalyzeState);
  });
})();
