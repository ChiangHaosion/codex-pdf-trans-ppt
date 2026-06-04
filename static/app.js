const pdfFile = document.querySelector("#pdfFile");
const fileName = document.querySelector("#fileName");
const sampleSource = document.querySelector("#sampleSource");
const customSkip = document.querySelector("#customSkip");
const skipModeInputs = document.querySelectorAll('input[name="skip_mode"]');
const convertForm = document.querySelector("#convertForm");
const submitButton = document.querySelector("#submitButton");
const precheckButton = document.querySelector("#precheckButton");
const precheckPanel = document.querySelector("#precheckPanel");
const previewPanel = document.querySelector(".preview-panel[data-job-status]");
const confirmForms = document.querySelectorAll("form[data-confirm]");

function element(tag, className, text) {
  const node = document.createElement(tag);
  if (className) {
    node.className = className;
  }
  if (text !== undefined) {
    node.textContent = text;
  }
  return node;
}

function resetPrecheck() {
  if (!precheckPanel) {
    return;
  }
  precheckPanel.hidden = true;
  precheckPanel.className = "precheck-panel";
  precheckPanel.replaceChildren();
}

function refreshCustomSkip() {
  const selected = document.querySelector('input[name="skip_mode"]:checked');
  customSkip?.classList.toggle("is-visible", selected?.value === "custom");
}

pdfFile?.addEventListener("change", () => {
  const selectedFile = pdfFile.files?.[0];
  fileName.textContent = selectedFile ? selectedFile.name : "选择 PDF 文件";
  if (selectedFile && sampleSource) {
    sampleSource.checked = false;
  }
  resetPrecheck();
});

sampleSource?.addEventListener("change", () => {
  if (sampleSource.checked && pdfFile) {
    pdfFile.value = "";
    fileName.textContent = "选择 PDF 文件";
  }
  resetPrecheck();
});

skipModeInputs.forEach((input) =>
  input.addEventListener("change", () => {
    refreshCustomSkip();
    resetPrecheck();
  }),
);
refreshCustomSkip();

document.querySelectorAll("#layout, #profile, #answerMode, #skipFromPage").forEach((input) => {
  input.addEventListener("change", resetPrecheck);
});

convertForm?.addEventListener("submit", () => {
  submitButton.classList.add("is-loading");
  submitButton.textContent = "生成中...";
  submitButton.disabled = true;
});

function addPrecheckMetric(metrics, label, value) {
  const item = element("div");
  item.append(element("dt", "", label), element("dd", "", String(value)));
  metrics.append(item);
}

function addPrecheckList(title, items) {
  if (!precheckPanel || !items.length) {
    return;
  }
  const list = element("div", "precheck-list");
  list.append(element("h4", "", title));
  items.slice(0, 6).forEach((item) => list.append(element("span", "", item)));
  if (items.length > 6) {
    list.append(element("span", "", `另有 ${items.length - 6} 项`));
  }
  precheckPanel.append(list);
}

function showPrecheckError(message) {
  if (!precheckPanel) {
    return;
  }
  precheckPanel.hidden = false;
  precheckPanel.className = "precheck-panel is-warn";
  precheckPanel.replaceChildren(element("p", "notice error", message));
}

function renderPrecheck(payload) {
  if (!precheckPanel) {
    return;
  }
  const result = payload.precheck;
  const hasWarnings = result.warnings.length > 0 || result.unmatched_topics > 0 || result.unused_answers > 0;
  precheckPanel.hidden = false;
  precheckPanel.className = `precheck-panel ${hasWarnings ? "is-warn" : "is-good"}`;
  precheckPanel.replaceChildren();

  const title = element("div", "precheck-title");
  title.append(element("h3", "", "预检查结果"), element("span", "", payload.original_name || ""));
  precheckPanel.append(title);

  const metrics = element("dl", "precheck-metrics");
  addPrecheckMetric(metrics, "答案页", result.answer_start_page);
  addPrecheckMetric(metrics, "题目", result.topic_count);
  addPrecheckMetric(metrics, "答案段", result.answer_count);
  addPrecheckMetric(metrics, "匹配", `${result.matched_answers}/${result.topic_count}`);
  precheckPanel.append(metrics);

  const tags = element("div", "precheck-tags");
  tags.append(element("span", "", `高 ${result.confidence_counts.high || 0}`));
  tags.append(element("span", "", `中 ${result.confidence_counts.medium || 0}`));
  tags.append(element("span", "", `低 ${result.confidence_counts.low || 0}`));
  tags.append(element("span", "", `未匹配 ${result.confidence_counts.none || 0}`));
  precheckPanel.append(tags);

  addPrecheckList("提示", result.warnings);
  addPrecheckList(
    "未匹配题目",
    result.unmatched_items.map((item) => item.title),
  );
  addPrecheckList(
    "未使用答案",
    result.unused_answer_items.map((item) => `${item.category || "答案"} ${item.number || ""} ${item.title}`.trim()),
  );
}

precheckButton?.addEventListener("click", async () => {
  if (!convertForm || !precheckButton.dataset.precheckUrl) {
    return;
  }
  precheckButton.classList.add("is-loading");
  precheckButton.textContent = "检查中";
  precheckButton.disabled = true;

  try {
    const response = await fetch(precheckButton.dataset.precheckUrl, {
      method: "POST",
      body: new FormData(convertForm),
      headers: { Accept: "application/json" },
    });
    const payload = await response.json();
    if (!response.ok || !payload.ok) {
      showPrecheckError(payload.error || "预检查失败");
      return;
    }
    renderPrecheck(payload);
  } catch {
    showPrecheckError("预检查失败，请重试");
  } finally {
    precheckButton.classList.remove("is-loading");
    precheckButton.textContent = "预检查";
    precheckButton.disabled = false;
  }
});

confirmForms.forEach((form) => {
  form.addEventListener("submit", (event) => {
    const message = form.dataset.confirm || "确认执行？";
    if (!window.confirm(message)) {
      event.preventDefault();
    }
  });
});

function pollJobStatus() {
  if (!previewPanel) {
    return;
  }
  const status = previewPanel.dataset.jobStatus;
  const statusUrl = previewPanel.dataset.statusUrl;
  const jobUrl = previewPanel.dataset.jobUrl;
  if (!["queued", "running"].includes(status) || !statusUrl || !jobUrl) {
    return;
  }

  window.setTimeout(async () => {
    try {
      const response = await fetch(statusUrl, { headers: { Accept: "application/json" } });
      if (!response.ok) {
        return;
      }
      const data = await response.json();
      if (["done", "failed"].includes(data.status)) {
        window.location.href = jobUrl;
        return;
      }
      previewPanel.dataset.jobStatus = data.status;
      pollJobStatus();
    } catch {
      pollJobStatus();
    }
  }, 1500);
}

pollJobStatus();
