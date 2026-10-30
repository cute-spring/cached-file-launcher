# 文档本地缓存下载器 Demo

这是一个零依赖的纯前端 Demo：`index.html`、原生 JavaScript、CSS 和 `sw.js`。它只调用已经存在的文件下载 API；首次下载或手动预缓存时，Service Worker 将 API 返回的**原始 HTTP Response**（含 DOCX、XLSX、PDF 等二进制响应体及下载头）写入 Cache Storage。已缓存文件的后续下载优先返回本地缓存，同时在后台尝试重新验证。

它适用于内部系统中重复下载 `.docx`、`.xlsx`、`.pdf` 等附件的场景，用以降低重复下载等待和后端带宽。不提供、也不尝试文档预览、解析、转换、编辑，或强制拉起本机应用。

“打开”按钮在纯网页模式下会说明浏览器限制，不会重复下载。可选的 `chrome-extension/` 扩展使用 Chrome 下载记录中的下载项 ID 请求系统打开**已完成的原文件**，不从 Cache Storage 重建下载文件，也不请求下载 API。它需要另行安装，并依赖操作系统已经关联 Word、WPS、Excel 或 PDF 阅读器。Cache Storage 中有副本却未曾下载到常规目录时，无法直接打开已有本机文件。

## 工程结构

```text
office-document-cache-demo/
├── index.html
├── sw.js
├── chrome-extension/
│   ├── manifest.json
│   ├── content.js
│   └── background.js
├── demo-files/
│   ├── star-young-writer-photography-role.docx
│   ├── service-record-template.xlsx
│   ├── compliance-notice.docx
│   └── cache-usage-guide.pdf
├── demo-server.py
├── tests/
│   ├── generate-sample-office.py
│   └── extension.test.cjs
└── README.md
```

## 架构与请求流程

```text
用户点击下载 / 预缓存
        │
        ▼
页面通过 postMessage 请求 Service Worker
        │
        ├── 未缓存：GET 既有下载 API → 校验 Office 响应 → 原始 Response 写入 Cache Storage
        │                                                    │
        │                                                    └→ 下载操作再请求同一 URL
        └── 已缓存：立即以附件响应下载 ← Cache Storage；后台条件请求检查更新
```

预缓存的进度由 Service Worker 使用流式 `ReadableStream.tee()` 的计数分支得到，另一个分支直接进入 Cache Storage；没有将完整文件读入 JS Blob、Base64、Data URL 或字符串。没有 `localStorage`、`sessionStorage` 或 IndexedDB 用于保存文件。

## 首先配置真实下载 API

修改 `index.html` 顶部的 `OFFICE_FILES`，仅替换为现有 API：

```js
{
  id: "policy-001",
  title: "客户服务操作指引",
  url: "/existing-download-api/files/policy-001",
  expectedFileName: "customer-service-guide.docx",
  type: "docx",
  expectedMimeTypes: [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  ]
}
```

同时让 `sw.js` 顶部的 `SAME_ORIGIN_API_PATH_PREFIXES` 保持与这些同源 URL 的路径前缀一致。这样 Service Worker 只会截获受允许的下载请求，不会误缓存网站的其他 GET 请求。

若 API 是跨域地址，还需把其 origin 明确填入 `ALLOWED_CROSS_ORIGIN_DOWNLOAD_ORIGINS`。跨域方案的 CORS 与认证要求见下文；建议优先使用同源部署，最容易获得稳定的附件下载和缓存命中行为。

## 后端既有 API 的前置条件

无需新建后端接口，但现有下载 API 应返回：

```http
Content-Disposition: attachment; filename="example.docx"
Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document
```

Excel 示例 MIME：

```http
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
```

PDF 示例 MIME：

```http
Content-Type: application/pdf
```

建议同时返回：

```http
Content-Length: <文件大小>
ETag: "<资源版本标识>"
Last-Modified: <HTTP 日期>
Cache-Control: no-cache
X-Content-Type-Options: nosniff
```

- `Content-Disposition: attachment` 决定浏览器应按附件下载，而不是在页面内预览。
- Demo 有意不设置 HTML 的 `download` 属性：部分 Chromium 下载路径会绕过 Service Worker。附件下载完全依赖上述 `Content-Disposition`，以保证已缓存响应仍可被 Service Worker 返回。
- `Content-Length` 可让预缓存显示百分比；没有它时页面显示不确定进度。
- `ETag` 是最优先的更新校验机制；`Last-Modified` 是备选。
- 不应使用 `Cache-Control: no-store`，除非安全策略明确禁止任何客户端缓存；它与本 Demo 的 Cache Storage 目的冲突。
- API 不得将未登录跳转页、HTML 错误页或 JSON 错误响应伪装成附件。Demo 会拒绝缓存常见 HTML/JSON/text 响应以及没有附件/受支持文档特征的响应。

### 条件请求

后端应支持：

```http
If-None-Match: "<ETag>"
```

或：

```http
If-Modified-Since: <HTTP 日期>
```

未变化时返回 `304 Not Modified`；变化时返回带完整二进制主体的 `200 OK`。Service Worker 的策略是**缓存优先 + 后台重新验证（stale-while-revalidate）**：已缓存附件先立即下载，验证结果决定下一次下载是否使用新版本。因此它不承诺每一次点击都绝对拿到服务器最新内容；网络不可用或后台验证尚未完成时仍会使用可用的本地版本。

若 API 没有 `ETag` 和 `Last-Modified`，Demo 以 `FALLBACK_REVALIDATE_TTL_MS`（默认 15 分钟）作为降级策略，到期后“检查更新”会重新下载并覆盖缓存；也可点单项“刷新缓存”。这不能可靠判断内容是否真正变化，UI 会明确提示该限制。

### 跨域与认证

跨域 API 必须正确配置 CORS，至少：

```http
Access-Control-Allow-Origin: https://your-frontend-origin.example
Access-Control-Allow-Methods: GET, HEAD, OPTIONS
Access-Control-Allow-Headers: If-None-Match, If-Modified-Since, Authorization, Content-Type
Access-Control-Expose-Headers: Content-Disposition, Content-Type, Content-Length, ETag, Last-Modified
```

Cookie 会话认证还应有：

```http
Access-Control-Allow-Credentials: true
```

带凭证时不能与 `Access-Control-Allow-Origin: *` 组合。当前 Demo 默认以 `same-origin` 凭证策略创建预缓存/更新检查请求；如果真实跨域 API 依赖 Cookie 或 Authorization，必须由现有 API 的 CORS、认证策略及经安全评审的前端请求策略共同支持。普通附件链接的认证行为也必须在目标浏览器中实测。

短期签名 URL、一次性 Token、频繁变化的 query 参数会改变 Cache Storage 键，降低命中率甚至令旧缓存不可用。应让既有网关/下载服务提供稳定资源标识或可靠版本机制。页面只显示 URL 的 origin 和 pathname，不显示 query 参数，避免把 Token 暴露在 UI。

## 本地运行与部署

Service Worker 只能在 HTTPS 或 `localhost` 下运行，不能用 `file://` 直接打开 `index.html`。

1. 先替换 `OFFICE_FILES` 和 `sw.js` 的下载白名单为可访问的既有 API。
2. 在该目录用任意静态文件服务器启动。例如 Mac 已有 Python 3 时：

   ```bash
   cd office-document-cache-demo
   python3 -m http.server 8080
   ```

3. 打开 `http://localhost:8080`。
4. 生产部署时，让 `index.html` 和 `sw.js` 位于同一 HTTPS 站点和预期 scope；下载 API 优先同源部署。

普通静态服务器会直接暴露 `demo-files/` 目录，含用户提供的演示文档；仅在隔离的本机测试环境使用。正式部署时移除 `demo-files/`、`demo-server.py`、`tests/` 和 `chrome-extension/`（扩展应单独审核部署）。本地 fixture 会阻止经静态路径直接读取该文档。

产品代码不新增、修改或依赖新的后端业务 API。仅用普通静态服务器时，示例 `/existing-download-api/files/...` 不存在，预缓存会失败；替换为真实的现有 API 后即可验证。

为复现**本地演示测试**，运行 `python3 demo-server.py` 并打开 `http://127.0.0.1:8080/index.html`。这个仅供本机测试的 fixture 将 DOCX、XLSX 和 PDF 文件映射到四个配置 URL，支持 ETag / `304`；不属于产品后端，也不得生产部署。四个文件均可用于“预缓存全部文件”测试。若需重新生成 XLSX、PDF 和第二份 DOCX，运行 `python3 tests/generate-sample-office.py`。扩展逻辑测试可用 `node --test tests/extension.test.cjs` 执行。

### 可选的 Chrome / Edge 本机打开扩展

纯网页没有“打开下载目录中的已有文件”的权限。要实现这个新增要求，需要在受管理的 Chrome / Edge Profile 中部署本目录下的 MV3 扩展：

1. 在本地测试时，将 `chrome-extension/` 作为**未打包扩展**加载。此扩展申请 `downloads` 和 `downloads.open` 权限，前者可读取当前浏览器 Profile 的下载记录，后者可请求系统打开下载项；只应在可信测试 Profile 中启用。真实环境建议由企业统一审核、签名及部署。
2. 示例 manifest 只匹配 `http://localhost:8080/index.html` 与 `http://127.0.0.1:8080/index.html`。真实部署时，将 manifest 中的 `matches` 和 `host_permissions` 改为实际 HTTPS 页面 origin，同时修改 `background.js` 的允许 origin 和下载 API 路径，并重新加载扩展。
3. 安装扩展后刷新页面。先点击“下载”，等 Chrome 下载项完成，再点击对应“打开”。扩展根据当前页面的文件 ID 寻找**相同下载 API URL** 的最新已完成 DOCX、XLSX 或 PDF 下载项，并调用 `chrome.downloads.open(downloadId)`。浏览器或系统可能弹出安全确认，具体应用由文件关联决定。
   页面顶部会显示是否检测到扩展，并明确列出需要授予的 `downloads`、`downloads.open` 权限及启用步骤。此状态仅说明扩展内容脚本已加载，不等于系统已经成功打开文件。
4. 如果文件移走、下载项被清理、下载尚未完成，或 Chrome 拒绝该用户手势，页面显示错误。预缓存只写 Cache Storage，因此单独预缓存不会产生可供扩展打开的下载项。

扩展拦截“打开”按钮的真人点击，校验页面 origin、API 路径和已下载文件类型。它不读取文档正文、不上传文件、不生成第二份下载，也不会主动获取访问令牌。扩展代码是可选增强；**单靠原始纯前端技术约束，不能实现这个新增的本机一键打开要求**。用户退出、切换身份或失去权限时，除了清理站点缓存，真实系统还须考虑下载目录中已落盘文件的管理；浏览器缓存清理不会删除这些文件。

## 测试清单

1. **首次下载**：点击“下载”。页面先将 API 的有效文档响应缓存，再触发附件下载；状态应成为“已缓存”。
2. **再次命中缓存**：再次点击同一文件；在 DevTools Network 中可观察请求由 Service Worker 提供，同时后端收到后台条件验证（如网络可用）。
3. **手动预缓存**：点击单项“预缓存”或“预缓存全部文件”。应不向常规下载目录写入文件。
4. **检查更新**：点击“检查更新”。带 ETag 的 API 应收到 `If-None-Match`；未更新返回 `304`，更新后返回新 `200` 并替换缓存。
5. **模拟 304**：后端对相同 ETag 返回 `304 Not Modified`；页面显示继续使用本地版本。
6. **模拟文件更新**：改变后端文件并返回新 ETag/新 `200`；页面显示“发现新版本并已更新”，下一次下载为新内容。
7. **清空缓存**：点击“清空全部本地缓存”并确认；所有条目应恢复为“未缓存”。这不会删除电脑下载目录里的文件。
8. **不支持 Service Worker 的降级**：在不支持的环境中，状态显示“本地缓存不可用”；下载仍走普通 API 链接，不阻断用户。
9. **错误响应防护**：让 API 返回登录 HTML、JSON 或非 2xx；Demo 应报错且不将其写入 Cache Storage。
10. **本机打开**：在扩展已启用、系统已安装对应文档应用或 PDF 阅读器的测试 Profile 中，记录下载目录文件数；首次下载并完成后点击“打开”，确认本机默认程序打开**原文件**，下载目录文件数不增加，下载记录不新增、服务端没有新的完整文件 GET。再试文件移走/下载记录清理、仅预缓存未下载、没有扩展三种失败路径，确认提示明确且不产生重复下载。

## DevTools 检查方法

Chrome / Edge：打开 DevTools → **Application**。

- **Service Workers**：确认 `sw.js` 已 activated / controlled；首次安装若尚未接管，可刷新页面一次。
- **Cache Storage**：展开 `office-download-cache-v2`，检查下载 URL 及保留的 `Content-Disposition`、`Content-Type`、ETag 等响应头。
- **Network**：禁用缓存、切换离线或观察请求的 Service Worker 来源，用于验证命中与回退。

## 缓存版本升级

`sw.js` 顶部定义：

```js
const CACHE_VERSION = "v2";
const CACHE_NAME = `office-download-cache-${CACHE_VERSION}`;
```

升级缓存结构、校验逻辑或需要清理旧内容时，将 `v2` 改为新值（如 `v3`）。新 Service Worker 在 `activate` 中删除此前以 `office-download-cache-` 开头的旧缓存，并以 `clients.claim()` 尽快接管当前页面。因此版本号变更会清空旧文档缓存；请将此行为纳入发布说明。

## 安全边界与已知限制

- Cache Storage 只是当前浏览器 Profile、当前 origin 沙盒内的临时性能缓存，不替代企业文档授权、DRM、审计、密钥管理或加密文档库。
- 登出、切换用户、切换租户、权限撤销时，真实系统必须调用同等的 `CLEAR_CACHE` 清理逻辑；不要让多人共用同一个浏览器 Profile。
- 浏览器可能因清理站点数据、隐私模式、存储压力或企业策略删除 Cache Storage；页面会自然回退为重新请求既有 API。
- 页面不会读取、扫描或访问用户电脑磁盘；也不会在浏览器内预览文档。
- 浏览器下载后是否自动打开 Word、Excel 或 PDF 阅读器完全取决于浏览器设置、企业策略及操作系统默认文件关联；前端只能发起附件下载，不能强制调用本地 Native App。
- 可选扩展可请求 Chrome 打开已有的已完成下载项，仍受用户手势、扩展安装权限、浏览器策略和操作系统文件关联限制；API 成功不等于可以证明目标应用已经显示文档。
- Cache Storage 保留 Response 头以实现下载；应用本身不得在 URL、日志或 UI 中泄漏 Token、签名参数或文件内容。
- 对跨域、SSO、下载拦截、安全软件及各浏览器的原生附件行为，必须在企业真实环境测试。最可靠的部署形式仍是前端与下载 API 同源。
