(function () {
  const defaultApiUrl = "https://cleardrop.wit-software.com/analyze";
  const analyzeButton = document.getElementById("analyze");
  const apiInput = document.getElementById("api-url");
  const statusEl = document.getElementById("status");
  const resultEl = document.getElementById("result");

  function setStatus(message, isError) {
    statusEl.textContent = message;
    statusEl.classList.toggle("muted", !isError);
    statusEl.style.color = isError ? "#b00020" : "";
  }

  function setResult(value) {
    resultEl.textContent = value || "";
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
        return;
      }

      let output = text;
      try {
        output = JSON.stringify(JSON.parse(text), null, 2);
      } catch (error) {
        // keep raw text
      }

      setStatus("Analysis complete.");
      setResult(output);
    } catch (error) {
      setStatus("Network error sending to API.", true);
      setResult(String(error));
    }
  }

  Office.onReady(() => {
    loadApiUrl();
    updateAnalyzeState();
    analyzeButton.addEventListener("click", analyzeEmail);
    apiInput.addEventListener("input", updateAnalyzeState);
  });
})();
