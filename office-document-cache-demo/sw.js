/* 文档本地缓存下载器：仅缓存明确允许的既有下载 API 的成功附件响应。 */
const CACHE_VERSION = "v2";
const CACHE_NAME = `office-download-cache-${CACHE_VERSION}`;
const CACHE_STORED_AT_HEADER = "X-Office-Demo-Cached-At";

// 必须与 index.html 配置区匹配。跨域 API 请显式加入其 origin，避免把任意请求当成文档下载请求缓存。
const SAME_ORIGIN_API_PATH_PREFIXES = ["/existing-download-api/files/"];
const ALLOWED_CROSS_ORIGIN_DOWNLOAD_ORIGINS = []; // 例如：["https://download-api.example.com"]

const inFlight = new Map();
const revalidationInFlight = new Map();

self.addEventListener("install", () => {
  // 不预缓存 Office 二进制文件；文件只会在用户下载或明确预缓存时进入 Cache Storage。
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(names.filter((name) => name.startsWith("office-download-cache-") && name !== CACHE_NAME).map((name) => caches.delete(name)));
    await self.clients.claim();
  })());
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET" || !isAllowedDownloadUrl(event.request.url)) return;
  event.respondWith(handleDownloadFetch(event));
});

function isAllowedDownloadUrl(rawUrl) {
  // 页面配置可以安全地使用同源相对路径（如 /existing-download-api/files/policy-001）。
  const url = new URL(rawUrl, self.location.origin);
  if (url.origin === self.location.origin) return SAME_ORIGIN_API_PATH_PREFIXES.some((prefix) => url.pathname.startsWith(prefix));
  return ALLOWED_CROSS_ORIGIN_DOWNLOAD_ORIGINS.includes(url.origin);
}
function cacheKey(url) { return new Request(url, { method: "GET" }); }
function responseInit(response, extraHeaders = {}) {
  const headers = new Headers(response.headers);
  Object.entries(extraHeaders).forEach(([key, value]) => headers.set(key, value));
  return { status: response.status, statusText: response.statusText, headers };
}
function isValidDocumentResponse(response, expectedMimeTypes = []) {
  if (!response || !response.ok || response.type === "opaque") return false;
  const contentType = (response.headers.get("Content-Type") || "").toLowerCase();
  const disposition = (response.headers.get("Content-Disposition") || "").toLowerCase();
  if (contentType.includes("text/html") || contentType.includes("application/json") || contentType.includes("text/plain")) return false;
  if (disposition && !disposition.includes("attachment")) return false;
  const knownDocument = contentType.includes("application/vnd.openxmlformats-officedocument") || contentType.includes("application/msword") || contentType.includes("application/vnd.ms-excel") || contentType.includes("application/pdf");
  const expected = expectedMimeTypes.length === 0 || expectedMimeTypes.some((type) => contentType.includes(type.toLowerCase()));
  return expected && (Boolean(disposition) || knownDocument);
}
function normalizeFile(file) {
  if (!file || typeof file.id !== "string" || typeof file.url !== "string" || !isAllowedDownloadUrl(file.url)) throw new Error("文件配置无效或 URL 不在 Service Worker 下载 API 白名单内。");
  return file;
}
async function cachedResponse(url) { return (await caches.open(CACHE_NAME)).match(cacheKey(url)); }
async function cacheResponse(url, response) {
  const cache = await caches.open(CACHE_NAME);
  const responseForCache = new Response(response.body, responseInit(response, { [CACHE_STORED_AT_HEADER]: new Date().toISOString() }));
  await cache.put(cacheKey(url), responseForCache);
}
async function handleDownloadFetch(event) {
  const { request } = event;
  const cached = await cachedResponse(request.url);
  if (cached) {
    // stale-while-revalidate：附件下载立即用本地版本，网络验证在后台完成。
    event.waitUntil(checkForUpdate({ id: "", url: request.url, expectedMimeTypes: [] }, 0).catch(() => {}));
    return cached;
  }
  const running = inFlight.get(request.url);
  if (running) {
    await running.catch(() => {});
    const afterWait = await cachedResponse(request.url);
    if (afterWait) return afterWait;
  }
  try {
    const { response, cacheDone } = await fetchAndStartCaching(request.url, request, []);
    event.waitUntil(cacheDone.catch(() => {}));
    return response;
  } catch (error) {
    return new Response("文件当前不可下载（网络不可用或下载 API 返回了不可缓存的响应）。", { status: 503, headers: { "Content-Type": "text/plain; charset=utf-8" } });
  }
}
async function fetchAndStartCaching(url, request, expectedMimeTypes) {
  const existing = inFlight.get(url);
  if (existing) { await existing; const cached = await cachedResponse(url); if (cached) return { response: cached, cacheDone: Promise.resolve() }; }
  // 保留现有认证上下文；禁止浏览器 HTTP 缓存参与本次显式缓存写入。
  const networkRequest = new Request(request || url, { cache: "no-store" });
  const response = await fetch(networkRequest);
  // 认证跳转、权限失败或后端错误不应进入 Cache Storage，但仍须让浏览器收到原始网络响应。
  if (!isValidDocumentResponse(response, expectedMimeTypes)) return { response, cacheDone: Promise.resolve() };
  const cacheDone = cacheResponse(url, response.clone());
  inFlight.set(url, cacheDone);
  // catch the promise returned by finally too, otherwise a cache write failure can become an unhandled rejection.
  void cacheDone.finally(() => { if (inFlight.get(url) === cacheDone) inFlight.delete(url); }).catch(() => {});
  return { response, cacheDone };
}
async function precacheFile(file) {
  normalizeFile(file);
  const existing = inFlight.get(file.url);
  if (existing) { await existing; return { cached: true, message: "已等待进行中的下载并完成缓存" }; }
  const request = new Request(file.url, { method: "GET", cache: "no-store", credentials: "same-origin" });
  const response = await fetch(request);
  if (!isValidDocumentResponse(response, file.expectedMimeTypes || [])) throw new Error(`下载 API 返回的响应不是可缓存的文档附件（HTTP ${response.status}）。`);
  const cacheDone = streamResponseIntoCache(file, response);
  inFlight.set(file.url, cacheDone);
  try { await cacheDone; return { cached: true, message: "已缓存" }; }
  finally { if (inFlight.get(file.url) === cacheDone) inFlight.delete(file.url); }
}
async function streamResponseIntoCache(file, response) {
  // 使用 Streams tee：一支直接流入 Cache Storage，另一支仅计数并通知 UI，绝不把整个文件读入 Blob/Base64/字符串。
  if (!response.body) { await cacheResponse(file.url, response); return; }
  const [cacheStream, progressStream] = response.body.tee();
  const cachePromise = caches.open(CACHE_NAME).then((cache) => cache.put(cacheKey(file.url), new Response(cacheStream, responseInit(response, { [CACHE_STORED_AT_HEADER]: new Date().toISOString() }))));
  const total = Number(response.headers.get("Content-Length")) || 0;
  const reader = progressStream.getReader(); let received = 0;
  try {
    while (true) {
      const { done, value } = await reader.read(); if (done) break;
      received += value.byteLength;
      notifyClients({ type: "CACHE_PROGRESS", fileId: file.id, total, received, percent: total ? Math.min(100, Math.round((received / total) * 100)) : 0 });
    }
    await cachePromise;
  } catch (error) { try { await reader.cancel(error); } catch {} throw error; }
}
async function checkForUpdate(file, fallbackTtlMs = 0) {
  normalizeFile(file);
  const cached = await cachedResponse(file.url);
  if (!cached) return { cached: false, message: "本地尚未缓存；请先下载或预缓存。" };
  const etag = cached.headers.get("ETag"); const lastModified = cached.headers.get("Last-Modified");
  const storedAt = Date.parse(cached.headers.get(CACHE_STORED_AT_HEADER) || "") || 0;
  if (!etag && !lastModified) {
    if (!fallbackTtlMs || Date.now() - storedAt < fallbackTtlMs) return { cached: true, message: "后端未提供 ETag / Last-Modified；尚未到降级 TTL 检查时间，继续使用本地版本。" };
    return refreshFile(file, "后端未提供版本验证信息；已按 TTL 重新获取并替换本地缓存。");
  }
  const existing = revalidationInFlight.get(file.url);
  if (existing) return existing;
  const task = (async () => {
    const headers = new Headers(); if (etag) headers.set("If-None-Match", etag); else headers.set("If-Modified-Since", lastModified);
    let response;
    try { response = await fetch(new Request(file.url, { method: "GET", headers, cache: "no-store", credentials: "same-origin" })); }
    catch (error) { throw new Error(`网络验证失败：${error.message || error}`); }
    if (response.status === 304) return { cached: true, message: "服务器确认文件未变化，继续使用本地版本。" };
    if (!isValidDocumentResponse(response, file.expectedMimeTypes || [])) throw new Error(`更新检查返回不可用响应（HTTP ${response.status}），继续使用本地版本。`);
    await cacheResponse(file.url, response);
    if (file.id) notifyClients({ type: "CACHE_UPDATED", fileId: file.id });
    return { cached: true, updated: true, message: "发现新版本并已更新；下一次下载将使用新版本。" };
  })();
  revalidationInFlight.set(file.url, task);
  try { return await task; } finally { if (revalidationInFlight.get(file.url) === task) revalidationInFlight.delete(file.url); }
}
async function refreshFile(file, successMessage = "已重新获取并替换本地缓存。") {
  normalizeFile(file);
  const request = new Request(file.url, { method: "GET", cache: "no-store", credentials: "same-origin" });
  const response = await fetch(request);
  if (!isValidDocumentResponse(response, file.expectedMimeTypes || [])) throw new Error(`刷新缓存失败：下载 API 返回 HTTP ${response.status} 或非文档附件。`);
  await cacheResponse(file.url, response);
  if (file.id) notifyClients({ type: "CACHE_UPDATED", fileId: file.id });
  return { cached: true, updated: true, message: successMessage };
}
async function getStatuses(files) {
  const cache = await caches.open(CACHE_NAME);
  return Promise.all(files.map(async (file) => {
    normalizeFile(file); const response = await cache.match(cacheKey(file.url));
    return { id: file.id, cached: Boolean(response), cachedAt: response?.headers.get(CACHE_STORED_AT_HEADER) || null };
  }));
}
function reply(port, payload) { if (port) port.postMessage({ ok: true, ...payload }); }
function replyError(port, error) { if (port) port.postMessage({ ok: false, error: error?.message || String(error) }); }
async function notifyClients(message) { const clients = await self.clients.matchAll({ type: "window", includeUncontrolled: true }); clients.forEach((client) => client.postMessage(message)); }

self.addEventListener("message", (event) => {
  const { type } = event.data || {}; const port = event.ports[0];
  event.waitUntil((async () => {
    try {
      if (type === "GET_CACHE_STATUS") return reply(port, { files: await getStatuses(event.data.files || []) });
      if (type === "PRECACHE") return reply(port, await precacheFile(event.data.file));
      if (type === "REFRESH_CACHE") return reply(port, await refreshFile(event.data.file));
      if (type === "CHECK_UPDATE") return reply(port, await checkForUpdate(event.data.file, Number(event.data.fallbackTtlMs) || 0));
      if (type === "PRECACHE_ALL") {
        let successCount = 0; let failureCount = 0;
        for (const file of event.data.files || []) { try { await precacheFile(file); successCount++; } catch (error) { failureCount++; notifyClients({ type: "CACHE_ERROR", fileId: file.id, error: error.message }); } }
        return reply(port, { successCount, failureCount });
      }
      if (type === "CLEAR_CACHE") {
        // 等待已经开始的缓存写入，避免清理操作与同一 URL 的写入产生竞争。
        await Promise.all([...inFlight.values()].map((task) => task.catch(() => {})));
        await caches.delete(CACHE_NAME);
        return reply(port, { cleared: true });
      }
      if (type === "SKIP_WAITING") { await self.skipWaiting(); return reply(port, { skipped: true }); }
      throw new Error("不支持的 Service Worker 消息类型。");
    } catch (error) { replyError(port, error); }
  })());
});
