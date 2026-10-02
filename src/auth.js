// MSAL public-client auth with a DPAPI-encrypted on-disk token cache.
import os from "node:os";
import path from "node:path";
import fs from "node:fs";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import { PublicClientApplication } from "@azure/msal-node";
import {
  DataProtectionScope,
  PersistenceCachePlugin,
  PersistenceCreator,
} from "@azure/msal-node-extensions";

// Load .env from the project root (not the cwd: Claude launches this server from anywhere).
// Variables already set in the environment win over .env.
const ENV_FILE = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", ".env");
if (fs.existsSync(ENV_FILE)) process.loadEnvFile(ENV_FILE);

const required = (name) => {
  const value = process.env[name];
  if (!value) throw new Error(`Missing ${name}. Copy .env.example to .env and fill it in.`);
  return value;
};

const CLIENT_ID = required("GRAPH_CLIENT_ID");
const TENANT_ID = required("GRAPH_TENANT_ID");

// Add scopes here (or via GRAPH_SCOPES, space-separated) as new modules are added.
// offline_access is added by MSAL automatically; do not list it.
// Set to true once Team.ReadBasic.All, Channel.ReadBasic.All and ChannelMessage.Read.All are added
// in the app registration and admin-consented; requesting an unconsented scope breaks sign-in.
export const TEAMS_CHANNELS = false;

// Lists: items, getList and columns need Sites.ReadWrite.All (always requested).
// LISTS_MANAGE adds Sites.Manage.All, required by createList.
// LISTS_VIEWS adds the SharePoint Online API permission AllSites.Manage, required by getViews/createView
// (Graph has no views API, so these call SharePoint REST with a separate token).
// Flip a switch only after the matching permission is added in the app registration.
export const LISTS_MANAGE = true;
export const LISTS_VIEWS = true;

export const SHAREPOINT_ROOT_HOST = LISTS_VIEWS ? required("SHAREPOINT_ROOT_HOST") : process.env.SHAREPOINT_ROOT_HOST;
export const spScope = (host) => `https://${host}/AllSites.Manage`;

const DEFAULT_SCOPES = [
  "Notes.ReadWrite.All",
  "Mail.ReadWrite",
  "Chat.Read",
  "Sites.ReadWrite.All",
  ...(TEAMS_CHANNELS ? ["Team.ReadBasic.All", "Channel.ReadBasic.All", "ChannelMessage.Read.All"] : []),
  ...(LISTS_MANAGE ? ["Sites.Manage.All"] : []),
];
export const SCOPES = (process.env.GRAPH_SCOPES ?? DEFAULT_SCOPES.join(" ")).split(/\s+/).filter(Boolean);

// A token is only valid for one resource, so SharePoint REST scopes are consented at login but
// requested separately (per host) when used.
const EXTRA_CONSENT_SCOPES = LISTS_VIEWS ? [spScope(SHAREPOINT_ROOT_HOST)] : [];

const CACHE_PATH =
  process.env.GRAPH_TOKEN_CACHE ?? path.join(os.homedir(), ".graph-mcp", "token_cache.bin");

let pcaPromise;
function getApp() {
  pcaPromise ??= (async () => {
    fs.mkdirSync(path.dirname(CACHE_PATH), { recursive: true });
    const persistence = await PersistenceCreator.createPersistence({
      cachePath: CACHE_PATH,
      dataProtectionScope: DataProtectionScope.CurrentUser, // Windows DPAPI, tied to your user
      serviceName: "graph-mcp",
      accountName: "msal-cache",
      usePlaintextFileOnLinux: true,
    });
    return new PublicClientApplication({
      auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}` },
      cache: { cachePlugin: new PersistenceCachePlugin(persistence) },
      // MSAL logs go to stderr only; stdout is reserved for the MCP protocol.
    });
  })();
  return pcaPromise;
}

function openBrowser(url) {
  // `start` treats & as a separator, so spawn via cmd with the URL quoted.
  spawn("cmd", ["/c", "start", '""', `"${url}"`], { windowsHide: true, detached: true, shell: true, stdio: "ignore" }).unref();
}

/** Returns an access token. Silent from cache; interactive (browser) only when asked. */
export async function getToken({ interactive = false, scopes = SCOPES } = {}) {
  const app = await getApp();
  const accounts = await app.getTokenCache().getAllAccounts();
  if (accounts.length) {
    try {
      const r = await app.acquireTokenSilent({ scopes, account: accounts[0] });
      if (r?.accessToken) return r.accessToken;
    } catch {
      // fall through to interactive / error
    }
  }
  if (!interactive) {
    throw new Error("Not signed in to Microsoft Graph. Run `npm run login` in the graph-mcp folder.");
  }
  const r = await app.acquireTokenInteractive({
    scopes,
    extraScopesToConsent: EXTRA_CONSENT_SCOPES,
    openBrowser: async (url) => openBrowser(url),
    successTemplate: "Signed in to graph-mcp. You can close this tab.",
    errorTemplate: "Sign-in failed: {error}. See the terminal for details.",
  });
  return r.accessToken;
}
