const API_PREFIX = "/existing-download-api/files/";
const SUPPORTED_DOCUMENT_EXTENSION = /\.(?:docx|xlsx|doc|xls|pdf)$/i;

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message?.type !== "OPEN_EXISTING_DOWNLOAD") return;
  try {
    const page = new URL(sender.url || "");
    const target = new URL(message.url || "");
    if (!(["http://localhost:8080", "http://127.0.0.1:8080"].includes(page.origin)
      && page.pathname === "/index.html" && target.origin === page.origin
      && target.pathname.startsWith(API_PREFIX) && !target.search && !target.hash)) {
      throw new Error("页面或下载地址不在允许范围内。");
    }
    chrome.downloads.search({ url: target.href, state: "complete" }, async (items) => {
      try {
        const item = items.filter((entry) => entry.url === target.href && SUPPORTED_DOCUMENT_EXTENSION.test(entry.filename || ""))
          .sort((a, b) => Date.parse(b.endTime || b.startTime || 0) - Date.parse(a.endTime || a.startTime || 0))[0];
        if (!item) throw new Error("没有找到已完成的对应下载；请先点击“下载”。");
        await chrome.downloads.open(item.id);
        sendResponse({ ok: true, downloadId: item.id });
      } catch (error) { sendResponse({ ok: false, error: error.message || String(error) }); }
    });
  } catch (error) { sendResponse({ ok: false, error: error.message || String(error) }); return; }
  return true;
});
