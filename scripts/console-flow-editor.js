// Power Automate flow editing from the browser DevTools console.
//
// Does the same thing as the extension (see src/background.ts and
// src/features/flow-editor/useFlowEditor.ts) without installing it:
// reuse the auth token of the open Power Automate tab and call the API directly.
//
// Usage: paste this whole file into the console of the make.powerautomate.com
// tab (the API's CORS policy allows that origin), then either:
//   a) fill in BASE / TOKEN below with values copied from DevTools > Network
//      (filter on "powerautomate/flows", copy the request host and the full
//      "authorization" header value), or
//   b) leave them empty and click around in the portal (open the flow list or
//      a flow) so the fetch/XHR hooks below capture them into window.PA.
//      Check what was captured with PA_status().
//
// Tokens expire after about an hour; on a 401, copy or capture a fresh one.
// Tokens are audience specific: the powerplatform.com token does not work
// against api.flow.microsoft.com and vice versa.

// New API: https://<encoded-env>.environment.api.powerplatform.com/
let BASE = "";
// Full header value, e.g. "Bearer eyJ0eXAi..."
let TOKEN = "";
// Legacy API (listing fallback + validation), e.g. https://<region>.api.flow.microsoft.com/
let LEGACY_BASE = "";
let LEGACY_TOKEN = "";

window.PA = window.PA || {};
window.PA.backups = window.PA.backups || {};

// Optional: capture tokens from the page's own requests (fetch and XHR).
(() => {
  if (window.__paHooked) {
    return;
  }
  window.__paHooked = true;

  const capture = (url, token) => {
    try {
      if (!url || !token) {
        return;
      }
      const absolute = new URL(url, location.href);
      const origin = absolute.origin + "/";
      if (/\.api\.powerplatform\.com$/i.test(absolute.hostname)) {
        if (/\/powerautomate\//i.test(absolute.pathname)) {
          window.PA.base = origin;
          window.PA.token = token;
        }
      } else if (/\.api\.flow\.microsoft\.com$/i.test(absolute.hostname)) {
        window.PA.legacyBase = origin;
        window.PA.legacyToken = token;
      }
    } catch (_) {
      // Never break the page's own requests.
    }
  };

  const originalFetch = window.fetch;
  window.fetch = function (input, init) {
    try {
      const url = typeof input === "string" ? input : input?.url;
      const headers = new Headers(
        init?.headers || (input instanceof Request ? input.headers : undefined)
      );
      capture(url, headers.get("authorization"));
    } catch (_) {}
    return originalFetch.apply(this, arguments);
  };

  const originalOpen = XMLHttpRequest.prototype.open;
  const originalSetHeader = XMLHttpRequest.prototype.setRequestHeader;
  XMLHttpRequest.prototype.open = function (method, url) {
    this.__paUrl = String(url);
    return originalOpen.apply(this, arguments);
  };
  XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
    if (String(name).toLowerCase() === "authorization") {
      capture(this.__paUrl, value);
    }
    return originalSetHeader.apply(this, arguments);
  };
})();

function PA_status() {
  return {
    base: BASE || window.PA.base || null,
    hasToken: Boolean(TOKEN || window.PA.token),
    legacyBase: LEGACY_BASE || window.PA.legacyBase || null,
    hasLegacyToken: Boolean(LEGACY_TOKEN || window.PA.legacyToken),
    ...getIdsFromLocation(),
  };
}

function hasApi(legacy) {
  return legacy
    ? Boolean((LEGACY_BASE || window.PA.legacyBase) && (LEGACY_TOKEN || window.PA.legacyToken))
    : Boolean((BASE || window.PA.base) && (TOKEN || window.PA.token));
}

function getApi(legacy) {
  const base = legacy ? LEGACY_BASE || window.PA.legacyBase : BASE || window.PA.base;
  const token = legacy ? LEGACY_TOKEN || window.PA.legacyToken : TOKEN || window.PA.token;
  if (!base || !token) {
    throw new Error(
      `Missing ${legacy ? "legacy " : ""}API base URL or token. ` +
        "Set them at the top of the script or interact with the portal so they get captured."
    );
  }
  return { base, token };
}

// `path` is relative to the API base, or an absolute URL (e.g. a nextLink,
// which already carries its own api-version).
async function request(legacy, path, method, data) {
  const { base, token } = getApi(legacy);
  const apiVersion = legacy ? "2016-11-01" : "1";
  let url = /^https?:\/\//i.test(path) ? path : base + path;
  if (!/[?&]api-version=/i.test(url)) {
    url += (url.includes("?") ? "&" : "?") + `api-version=${apiVersion}`;
  }

  const response = await fetch(url, {
    method,
    body: data ? JSON.stringify(data) : undefined,
    headers: {
      authorization: token,
      "Content-Type": "application/json",
    },
  });

  const text = await response.text();
  let body = null;
  if (text) {
    try {
      body = JSON.parse(text);
    } catch (_) {
      body = text;
    }
  }

  if (response.ok) {
    return body;
  }
  const message =
    (body && typeof body === "object" && (body.error?.message || body.message)) ||
    (typeof body === "string" ? body.slice(0, 500) : "") ||
    response.statusText;
  throw new Error(`${response.status}: ${message}`);
}

// Follows nextLink / @odata.nextLink until all pages are read.
async function requestAllPages(legacy, path) {
  const items = [];
  let next = path;
  while (next) {
    const page = await request(legacy, next, "GET");
    items.push(...(page?.value ?? []));
    next = page?.nextLink || page?.["@odata.nextLink"] || null;
  }
  return items;
}

// Flow ID / environment ID from the current tab URL:
// .../environments/{envId}/flows/{flowId}
function getIdsFromLocation() {
  const envId = /environments\/([a-zA-Z0-9\-]+)/i.exec(location.href)?.[1] ?? null;
  const flowId =
    /flows\/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i.exec(
      location.href
    )?.[1] ?? null;
  return { envId, flowId };
}

function legacyEnvPath(envId) {
  if (!envId) {
    throw new Error("Missing environment ID (open a page under /environments/{envId}/...)");
  }
  return `providers/Microsoft.ProcessSimple/environments/${envId}`;
}

function summarizeFlow(flow) {
  const p = flow.properties || {};
  return {
    id: flow.name || p.workflowEntityId || flow.id,
    displayName: p.displayName,
    state: p.state,
    createdTime: p.createdTime,
    lastModifiedTime: p.lastModifiedTime,
    creator: p.creator?.userId ?? p.creator?.objectId,
    triggerType: Object.values(p.definitionSummary?.triggers ?? [])?.[0]?.type,
  };
}

// Lists the flows of the current environment.
// options.source: "auto" (new API, then legacy), "new" or "legacy".
// options.filter (legacy only): "all" (personal + shared), "personal" or "team".
// Returns raw flow objects; console.table(flows.map(summarizeFlow)) for an overview.
async function listFlows({ source = "auto", filter = "all", envId = getIdsFromLocation().envId } = {}) {
  const errors = [];

  if ((source === "auto" || source === "new") && hasApi(false)) {
    try {
      return await requestAllPages(false, "powerautomate/flows");
    } catch (e) {
      errors.push(`new API: ${e.message}`);
      if (source === "new") {
        throw e;
      }
    }
  }

  if ((source === "auto" || source === "legacy") && hasApi(true)) {
    const path = `${legacyEnvPath(envId)}/flows`;
    const filters = filter === "all" ? ["personal", "team"] : [filter];
    const byId = new Map();
    for (const f of filters) {
      const flows = await requestAllPages(true, `${path}?$filter=search('${f}')`);
      for (const flow of flows) {
        byId.set(flow.name, flow);
      }
    }
    return [...byId.values()];
  }

  throw new Error(
    "Could not list flows. " +
      (errors.length ? errors.join("; ") + ". " : "") +
      "Open the 'My flows' page so the legacy token gets captured, or set LEGACY_BASE / LEGACY_TOKEN."
  );
}

// GET a single flow including its definition. Uses the new API when available,
// otherwise the legacy one. The first fetched version of every flow is kept in
// window.PA.backups[flowId] so it can be restored with restoreFlow(flowId).
async function getFlow(flowId = getIdsFromLocation().flowId, { envId = getIdsFromLocation().envId } = {}) {
  if (!flowId) {
    throw new Error("Missing flow ID");
  }
  let flow;
  try {
    flow = await request(false, `powerautomate/flows/${flowId}`, "GET");
  } catch (e) {
    if (!hasApi(true) || !envId) {
      throw e;
    }
    flow = await request(true, `${legacyEnvPath(envId)}/flows/${flowId}`, "GET");
  }
  if (!window.PA.backups[flowId]) {
    window.PA.backups[flowId] = structuredClone(flow);
  }
  return flow;
}

// PATCH powerautomate/flows/{flowId}
// Overwrites the live flow. definition and connectionReferences are both required.
// flowId is required on purpose so a flow is never overwritten by mistake.
async function saveFlow(flowId, { displayName, environment, definition, connectionReferences }) {
  if (!flowId) {
    throw new Error("Missing flow ID");
  }
  if (!definition) {
    throw new Error('Missing "definition" flow property');
  }
  if (!connectionReferences) {
    throw new Error('Missing "connectionReferences" flow property');
  }
  if (!window.PA.backups[flowId]) {
    await getFlow(flowId);
  }
  return request(false, `powerautomate/flows/${flowId}`, "PATCH", {
    properties: { displayName, environment, definition, connectionReferences },
  });
}

// Restores the version captured by the first getFlow(flowId) of this session.
async function restoreFlow(flowId) {
  const backup = window.PA.backups[flowId];
  if (!backup) {
    throw new Error(`No backup for flow ${flowId}`);
  }
  return saveFlow(flowId, backup.properties);
}

// POST .../checkFlowErrors and .../checkFlowWarnings on the legacy API
async function validateFlow(definition, ids = getIdsFromLocation()) {
  const path = `${legacyEnvPath(ids.envId)}/flows/${ids.flowId}`;
  const body = { properties: { definition } };
  const errors = await request(true, `${path}/checkFlowErrors`, "POST", body);
  const warnings = await request(true, `${path}/checkFlowWarnings`, "POST", body);
  return { errors, warnings };
}

// Fetches the definition of each { id, envId?, displayName?, ...extra } item.
// Returns [{ ...item, ...summary, definition, connectionReferences, error? }]
// in the same order; failed flows get an `error` instead of stopping the run.
async function exportFlows(items, { concurrency = 4, envId = getIdsFromLocation().envId } = {}) {
  const results = new Array(items.length);
  let index = 0;
  let done = 0;

  const worker = async () => {
    while (index < items.length) {
      const i = index++;
      const item = items[i];
      try {
        const flow = await getFlow(item.id, { envId: item.envId || envId });
        const summary = Object.fromEntries(
          Object.entries(summarizeFlow(flow)).filter(([, v]) => v !== undefined)
        );
        results[i] = {
          ...item,
          ...summary,
          id: item.id,
          definition: flow.properties?.definition,
          connectionReferences: flow.properties?.connectionReferences,
        };
      } catch (e) {
        results[i] = { ...item, error: e.message };
      }
      console.log(
        `[${++done}/${items.length}] ${item.displayName || item.id}` +
          (results[i].error ? ` - ERROR ${results[i].error}` : "")
      );
    }
  };

  await Promise.all(Array.from({ length: Math.min(concurrency, items.length) }, worker));
  const failed = results.filter((r) => r.error);
  if (failed.length) {
    console.warn(`${failed.length} flow(s) failed:`, failed.map((r) => `${r.displayName || r.id}: ${r.error}`));
  }
  return results;
}

// Fetches every flow of the environment with its definition.
async function exportAllFlows({ concurrency = 4, ...listOptions } = {}) {
  const flows = await listFlows(listOptions);
  return exportFlows(flows.map(summarizeFlow), { concurrency, envId: listOptions.envId });
}

// Minimal RFC 4180 CSV parser (quoted fields, "" escapes, CRLF) -> array of
// objects keyed by the header row.
function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;
  text = text.replace(/^\uFEFF/, "");

  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"' && text[i + 1] === '"') {
        field += '"';
        i++;
      } else if (c === '"') {
        quoted = false;
      } else {
        field += c;
      }
    } else if (c === '"') {
      quoted = true;
    } else if (c === ",") {
      row.push(field);
      field = "";
    } else if (c === "\n" || c === "\r") {
      if (c === "\r" && text[i + 1] === "\n") {
        i++;
      }
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += c;
    }
  }
  if (field || row.length) {
    row.push(field);
    rows.push(row);
  }

  const [header, ...data] = rows.filter((r) => r.some((v) => v.trim()));
  return data.map((r) => Object.fromEntries(header.map((h, i) => [h.trim(), (r[i] ?? "").trim()])));
}

// Fetches the flows listed in a Power Automate inventory CSV export
// (columns "Item ID", "Item name", "Environment ID"; other columns are kept
// under `inventory`). `csv` is the CSV text or already parsed rows.
async function exportFlowsFromCsv(csv, options = {}) {
  const rows = typeof csv === "string" ? parseCsv(csv) : csv;
  const items = rows
    .filter((r) => r["Item ID"])
    .map((r) => ({
      id: r["Item ID"],
      envId: r["Environment ID"] || undefined,
      displayName: r["Item name"],
      inventory: r,
    }));
  console.log(`Fetching ${items.length} flow(s) from the CSV...`);
  return exportFlows(items, options);
}

// Opens a file picker for the CSV. Call it directly from the console (not after
// an await) so the browser treats it as a user action.
function pickCsvFile() {
  return new Promise((resolve, reject) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".csv,text/csv";
    input.onchange = () => {
      const file = input.files?.[0];
      file ? file.text().then(resolve, reject) : reject(new Error("No file selected"));
    };
    input.click();
  });
}

// Pick the inventory CSV, fetch all its flows and download the result.
async function exportFlowsFromCsvFile(fileName = "flows.json", options = {}) {
  const text = await pickCsvFile();
  const results = await exportFlowsFromCsv(text, options);
  downloadJson(results, fileName);
  return results;
}

// Saves any value as a pretty-printed JSON file through the browser download.
function downloadJson(data, fileName = "flows.json") {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

// Examples:
//
//   PA_status();                                   // what has been captured
//
//   const flows = await listFlows();               // all flows of the current environment
//   console.table(flows.map(summarizeFlow));
//
//   const all = await exportAllFlows();            // every flow with its definition
//   downloadJson(all, "flows.json");
//
//   // Only the flows of an inventory CSV ("Item ID" / "Environment ID" columns):
//   const fromCsv = await exportFlowsFromCsvFile("flows.json");   // file picker + download
//   // or paste the CSV text directly:
//   // const fromCsv = await exportFlowsFromCsv(`...csv content...`);
//   // downloadJson(fromCsv, "flows.json");
//   console.table(fromCsv.map(({ id, displayName, state, error }) => ({ id, displayName, state, error })));
//
//   const { flowId } = getIdsFromLocation();
//   const flow = await getFlow(flowId);            // flow from the current URL
//   const props = flow.properties;
//   console.log(JSON.stringify({ definition: props.definition, connectionReferences: props.connectionReferences }, null, 2));
//
//   // edit props.definition, or replace it:
//   // props.definition = JSON.parse(`{ ...your flow definition... }`);
//
//   console.log(await validateFlow(props.definition));  // optional, needs the legacy token
//   console.log(await saveFlow(flowId, props));
//
//   // restore the original if something went wrong:
//   // await restoreFlow(flowId);
