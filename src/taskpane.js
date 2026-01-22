(function () {
  const defaultApiUrl = "https://cleardrop.wit-software.com/analyze";
  const uiVersion = "0.1.0";
  const analyzeButton = document.getElementById("analyze");
  const apiInput = document.getElementById("api-url");
  const statusEl = document.getElementById("status");
  const resultEl = document.getElementById("result");
  const gaugeEl = document.querySelector(".gauge");
  const gaugeVerdictEl = document.getElementById("gauge-verdict");
  const versionEl = document.getElementById("ui-version");

  function setStatus(message, isError) {
    statusEl.textContent = message;
    statusEl.classList.toggle("muted", !isError);
    statusEl.style.color = isError ? "#b00020" : "";
  }

  function setResult(value) {
    resultEl.textContent = value || "";
  }

  function updateGauge(score, verdict) {
    if (!gaugeEl || !gaugeVerdictEl) {
      return;
    }

    if (typeof score !== "number" || Number.isNaN(score)) {
      gaugeEl.style.setProperty("--gauge-rotate", "180deg");
      gaugeEl.dataset.state = "idle";
      gaugeVerdictEl.textContent = verdict || "Awaiting analysis";
      gaugeVerdictEl.classList.add("muted");
      return;
    }

    const clamped = Math.max(0, Math.min(100, score));
    const rotate = 180 - (clamped / 100) * 180;
    gaugeEl.style.setProperty("--gauge-rotate", `${rotate}deg`);
    gaugeVerdictEl.classList.remove("muted");

    const verdictText = verdict || (clamped < 35 ? "Safe" : clamped < 70 ? "Suspicious" : "Phishing");
    gaugeVerdictEl.textContent = verdictText;

    if (/safe/i.test(verdictText) || clamped < 35) {
      gaugeEl.dataset.state = "safe";
    } else if (/suspicious|warning/i.test(verdictText) || clamped < 70) {
      gaugeEl.dataset.state = "warning";
    } else {
      gaugeEl.dataset.state = "danger";
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
    updateGauge(null, "Collecting email...");

    let bodyText = "";
    try {
      bodyText = await getBodyText();
    } catch (error) {
      setStatus("Could not read email body.", true);
      return;
    }

    const payload = {
      subject: item.subject || "",
      from: item.from
        ? { name: item.from.displayName || "", address: item.from.emailAddress || "" }
        : null,
      to: getRecipients(item.to),
      cc: getRecipients(item.cc),
      itemId: item.itemId || "",
      internetMessageId: item.internetMessageId || "",
      bodyText,
      receivedDateTime: item.dateTimeCreated || "",
    };

    setStatus("Sending for analysis...");
    updateGauge(null, "Analyzing...");

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
        updateGauge(null, "Analysis failed");
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
      updateGauge(extractScore(parsed), extractVerdict(parsed));
    } catch (error) {
      setStatus("Network error sending to API.", true);
      setResult(String(error));
      updateGauge(null, "Network error");
    }
  }

  Office.onReady(() => {
    loadApiUrl();
    updateAnalyzeState();
    updateGauge(null, "Awaiting analysis");
    if (versionEl) {
      versionEl.textContent = `v${uiVersion}`;
    }
    analyzeButton.addEventListener("click", analyzeEmail);
    apiInput.addEventListener("input", updateAnalyzeState);
  });
})();
