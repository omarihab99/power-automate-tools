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
//   b) leave them empty and click around in the designer so the fetch hook
//      below captures them into window.PA.
//
// Tokens expire after about an hour; on a 401, copy or capture a fresh one.
// Tokens are audience specific: the powerplatform.com token does not work
// against api.flow.microsoft.com and vice versa.

// New API: https://<encoded-env>.environment.api.powerplatform.com/
let BASE = "";
// Full header value, e.g. "Bearer eyJ0eXAi..."
let TOKEN = "";
// Legacy API (validation only), e.g. https://<region>.api.flow.microsoft.com/
let LEGACY_BASE = "";
let LEGACY_TOKEN = "";

// Optional: capture tokens from the page's own requests.
(() => {
  if (window.__paFetchHooked) {
    return;
  }
  window.__paFetchHooked = true;
  window.PA = window.PA || {};

  const originalFetch = window.fetch;
  window.fetch = async (input, init) => {
    try {
      const url = typeof input === "string" ? input : input.url;
      const headers = new Headers(
        init?.headers || (input instanceof Request ? input.headers : undefined)
      );
      const token = headers.get("authorization");
      if (token) {
        const origin = new URL(url).origin + "/";
        if (/\.api\.powerplatform\.com\/powerautomate\/flows/i.test(url)) {
          window.PA.base = origin;
          window.PA.token = token;
        } else if (/\.api\.flow\.microsoft\.com\//i.test(url)) {
          window.PA.legacyBase = origin;
          window.PA.legacyToken = token;
        }
      }
    } catch (_) {
      // Never break the page's own requests.
    }
    return originalFetch(input, init);
  };
})();

function getApi(legacy) {
  const base = legacy ? LEGACY_BASE || window.PA.legacyBase : BASE || window.PA.base;
  const token = legacy ? LEGACY_TOKEN || window.PA.legacyToken : TOKEN || window.PA.token;
  if (!base || !token) {
    throw new Error(
      `Missing ${legacy ? "legacy " : ""}API base URL or token. ` +
        "Set them at the top of the script or interact with the designer so they get captured."
    );
  }
  return { base, token };
}

async function request(legacy, path, method, data) {
  const { base, token } = getApi(legacy);
  const apiVersion = legacy ? "2016-11-01" : "1";
  const url = base + path + (path.includes("?") ? "&" : "?") + `api-version=${apiVersion}`;

  const response = await fetch(url, {
    method,
    body: data ? JSON.stringify(data) : undefined,
    headers: {
      authorization: token,
      "Content-Type": "application/json",
    },
  });

  const body = await response.json();
  if (response.ok) {
    return body;
  }
  throw new Error(`${response.status}: ${body.error?.message}`);
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

// GET powerautomate/flows/{flowId}
async function getFlow(flowId = getIdsFromLocation().flowId) {
  const flow = await request(false, `powerautomate/flows/${flowId}`, "GET");
  // Keep the original so it can be restored with saveFlow(window.PA.backup.properties)
  window.PA.backup = window.PA.backup || structuredClone(flow);
  return flow;
}

// PATCH powerautomate/flows/{flowId}
// Overwrites the live flow. definition and connectionReferences are both required.
async function saveFlow(
  { displayName, environment, definition, connectionReferences },
  flowId = getIdsFromLocation().flowId
) {
  if (!definition) {
    throw new Error('Missing "definition" flow property');
  }
  if (!connectionReferences) {
    throw new Error('Missing "connectionReferences" flow property');
  }
  return request(false, `powerautomate/flows/${flowId}`, "PATCH", {
    properties: { displayName, environment, definition, connectionReferences },
  });
}

// POST .../checkFlowErrors and .../checkFlowWarnings on the legacy API
async function validateFlow(definition, ids = getIdsFromLocation()) {
  const path = `providers/Microsoft.ProcessSimple/environments/${ids.envId}/flows/${ids.flowId}`;
  const body = { properties: { definition } };
  const errors = await request(true, `${path}/checkFlowErrors`, "POST", body);
  const warnings = await request(true, `${path}/checkFlowWarnings`, "POST", body);
  return { errors, warnings };
}

// Example:
//
//   const flow = await getFlow();
//   const props = flow.properties;
//   console.log(JSON.stringify({ definition: props.definition, connectionReferences: props.connectionReferences }, null, 2));
//
//   // edit props.definition, or replace it:
//   // props.definition = JSON.parse(`{ ...your flow definition... }`);
//
//   console.log(await validateFlow(props.definition));  // optional, needs the legacy token
//   console.log(await saveFlow(props));
//
//   // restore the original if something went wrong:
//   // await saveFlow(window.PA.backup.properties);
