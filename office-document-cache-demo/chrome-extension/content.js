// DOM-only readiness signal for the page. Actual permissions are granted by Chrome at extension install.
function announceOpenHelper() {
  if (!document.documentElement) return;
  document.documentElement.dataset.officeOpenHelper = "ready";
  document.dispatchEvent(new Event("office-open-helper-ready"));
}
announceOpenHelper();
document.addEventListener("DOMContentLoaded", announceOpenHelper, { once: true });

// Only a physical click on the page's Open button can request a local open.
document.addEventListener("click", (event) => {
  const button = event.target.closest?.('button[data-action="open"][data-id]');
  if (!button || !event.isTrusted) return;
  const row = button.closest("[data-file-id]");
  if (!row || row.dataset.fileId !== button.dataset.id) return;
  const apiLabel = row.querySelector(".meta")?.textContent || "";
  const url = new URL(`/existing-download-api/files/${encodeURIComponent(button.dataset.id)}`, location.href).href;
  if (!apiLabel.includes(new URL(url).pathname)) return;

  // Stop the pure web fallback; the extension handles this click, without navigation.
  event.preventDefault();
  event.stopImmediatePropagation();
  button.disabled = true;
  chrome.runtime.sendMessage({ type: "OPEN_EXISTING_DOWNLOAD", url }, (result) => {
    button.disabled = false;
    const error = chrome.runtime.lastError?.message || result?.error;
    const notice = document.getElementById("notice");
    if (!notice) return;
    notice.textContent = error
      ? `无法打开已下载文件：${error}。请确认文件仍在下载目录，或从 Chrome 下载记录手动打开。`
      : "已请求系统用默认关联程序打开下载记录中的原文件；未产生新的下载。";
    notice.className = `notice show ${error ? "error" : "success"}`;
  });
}, true);
