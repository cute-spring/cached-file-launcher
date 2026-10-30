const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const path = require("node:path");

const url = "http://127.0.0.1:8080/existing-download-api/files/policy-001";
const page = "http://127.0.0.1:8080/index.html";
const source = fs.readFileSync(path.join(__dirname, "../chrome-extension/background.js"), "utf8");

function extension(downloads) {
  let listener;
  vm.runInNewContext(source, {
    URL,
    chrome: {
      runtime: { onMessage: { addListener(fn) { listener = fn; } } },
      downloads
    }
  });
  return (message, sender) => new Promise((resolve) => {
    const asyncReply = listener(message, sender, resolve);
    assert.ok(asyncReply === true || asyncReply === undefined);
  });
}

test("opens the completed original download without creating a new download", async () => {
  const calls = [];
  const send = extension({
    search(query, callback) {
      calls.push(["search", query.url]);
      callback([{ id: 17, url, state: "complete", filename: "/Downloads/star-young-writer-photography-role.docx", endTime: "2026-10-15T14:00:00Z" }]);
    },
    open(id) { calls.push(["open", id]); return Promise.resolve(); },
    download() { throw new Error("must not create another download"); }
  });
  const result = await send({ type: "OPEN_EXISTING_DOWNLOAD", url }, { url: page });
  assert.equal(result.ok, true);
  assert.deepEqual(calls, [["search", url], ["open", 17]]);
});

test("missing completed item reports an error and never opens anything", async () => {
  const send = extension({
    search(_query, callback) { callback([]); },
    open() { throw new Error("must not open"); }
  });
  const result = await send({ type: "OPEN_EXISTING_DOWNLOAD", url }, { url: page });
  assert.equal(result.ok, false);
  assert.match(result.error, /没有找到/);
});

test("browser refusal is shown as failure rather than claiming that the file opened", async () => {
  const send = extension({
    search(_query, callback) { callback([{ id: 12, url, filename: "/Downloads/guide.docx" }]); },
    open() { return Promise.reject(new Error("User gesture required")); }
  });
  const result = await send({ type: "OPEN_EXISTING_DOWNLOAD", url }, { url: page });
  assert.equal(result.ok, false);
  assert.match(result.error, /User gesture required/);
});

test("opens a completed PDF download using the same original-download path", async () => {
  const pdfUrl = "http://127.0.0.1:8080/existing-download-api/files/guide-004";
  const calls = [];
  const send = extension({
    search(query, callback) {
      calls.push(["search", query.url]);
      callback([{ id: 28, url: pdfUrl, state: "complete", filename: "/Downloads/cache-usage-guide.pdf", endTime: "2026-10-21T09:00:00Z" }]);
    },
    open(id) { calls.push(["open", id]); return Promise.resolve(); }
  });
  const result = await send({ type: "OPEN_EXISTING_DOWNLOAD", url: pdfUrl }, { url: page });
  assert.equal(result.ok, true);
  assert.deepEqual(calls, [["search", pdfUrl], ["open", 28]]);
});

test("rejects a different origin before inspecting download history", async () => {
  const send = extension({ search() { throw new Error("must not search"); } });
  // Synchronous rejections reply directly, instead of keeping a message port open.
  const result = await send({ type: "OPEN_EXISTING_DOWNLOAD", url: "https://example.com/private.docx" }, { url: page });
  assert.equal(result.ok, false);
  assert.match(result.error, /不在允许范围/);
});
