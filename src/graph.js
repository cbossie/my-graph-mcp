// Thin Microsoft Graph client shared by all tool modules.
import { getToken, spScope } from "./auth.js";

const BASE = "https://graph.microsoft.com/v1.0";

export async function request(method, path, { params, json, body, headers } = {}) {
  const url = new URL(path.startsWith("http") ? path : `${BASE}/${path.replace(/^\//, "")}`);
  for (const [k, v] of Object.entries(params ?? {})) url.searchParams.set(k, String(v));

  const hdrs = { Authorization: `Bearer ${await getToken()}`, ...headers };
  let payload = body;
  if (json !== undefined) {
    payload = JSON.stringify(json);
    hdrs["Content-Type"] = "application/json";
  }

  const res = await fetch(url, { method, headers: hdrs, body: payload });
  const text = await res.text();
  if (!res.ok) throw new Error(`Graph ${res.status} on ${method} ${path}: ${text.slice(0, 1000)}`);
  return (res.headers.get("content-type") ?? "").includes("json") ? JSON.parse(text) : text;
}

/**
 * SharePoint REST call (for things Graph can't do, e.g. list views).
 * `siteWebUrl` is the site's web URL (https://host/sites/x); the token is requested per host.
 */
export async function sharePointRequest(siteWebUrl, method, restPath, { json } = {}) {
  const host = new URL(siteWebUrl).host;
  const token = await getToken({ scopes: [spScope(host)] });
  // odata=verbose wraps results in { d: ... } but lets us send __metadata types on writes.
  const hdrs = { Authorization: `Bearer ${token}`, Accept: "application/json;odata=verbose" };
  let payload;
  if (json !== undefined) {
    payload = JSON.stringify(json);
    hdrs["Content-Type"] = "application/json;odata=verbose";
  }

  const res = await fetch(`${siteWebUrl.replace(/\/$/, "")}/_api/${restPath}`, { method, headers: hdrs, body: payload });
  const text = await res.text();
  if (!res.ok) throw new Error(`SharePoint ${res.status} on ${method} ${restPath}: ${text.slice(0, 1000)}`);
  if (!text) return undefined;
  const d = JSON.parse(text).d;
  return d?.results ?? d;
}
