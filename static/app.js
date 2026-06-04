const pdfFile = document.querySelector("#pdfFile");
const fileName = document.querySelector("#fileName");
const sampleSource = document.querySelector("#sampleSource");
const customSkip = document.querySelector("#customSkip");
const skipModeInputs = document.querySelectorAll('input[name="skip_mode"]');
const convertForm = document.querySelector("#convertForm");
const submitButton = document.querySelector("#submitButton");
const previewPanel = document.querySelector(".preview-panel[data-job-status]");
const confirmForms = document.querySelectorAll("form[data-confirm]");

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
});

sampleSource?.addEventListener("change", () => {
  if (sampleSource.checked && pdfFile) {
    pdfFile.value = "";
    fileName.textContent = "选择 PDF 文件";
  }
});

skipModeInputs.forEach((input) => input.addEventListener("change", refreshCustomSkip));
refreshCustomSkip();

convertForm?.addEventListener("submit", () => {
  submitButton.classList.add("is-loading");
  submitButton.textContent = "生成中...";
  submitButton.disabled = true;
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
