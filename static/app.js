const pdfFile = document.querySelector("#pdfFile");
const fileName = document.querySelector("#fileName");
const sampleSource = document.querySelector("#sampleSource");
const customSkip = document.querySelector("#customSkip");
const skipModeInputs = document.querySelectorAll('input[name="skip_mode"]');
const convertForm = document.querySelector("#convertForm");
const submitButton = document.querySelector("#submitButton");

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
