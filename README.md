# graph-mcp

A local [MCP](https://modelcontextprotocol.io) server for Microsoft Graph, written in Node. It runs on your machine, signs in as you with your own Entra app registration, and gives an MCP client (such as Claude) a set of Microsoft 365 tools. It is deliberately conservative: no tool sends mail or Teams messages, and OneNote has no delete or overwrite tools.

## Tools

| Area | Tools |
|---|---|
| **OneNote** | `onenote_list_notebooks`, `onenote_list_section_groups`, `onenote_list_sections`, `onenote_find_by_path`, `onenote_create_section`, `onenote_list_pages`, `onenote_get_page`, `onenote_create_page`, `onenote_append_to_page` |
| **Mail** | `mail_create_draft` (drafts only, never sends) |
| **Teams** | `teams_list_chats`, `teams_get_chat_messages` (read-only). Optional channel tools: `teams_list_teams`, `teams_list_channels`, `teams_get_channel_messages` |
| **Lists** (Microsoft Lists / SharePoint) | `lists_get_list`, `lists_get_columns`, `lists_create_item`, `lists_get_item`, `lists_update_item`, `lists_delete_item` (requires `confirm: true`), `lists_query_items`. Optional: `lists_create_list`, `lists_add_column`, `lists_update_column`, `lists_get_views`, `lists_create_view` |

Lists default to your personal OneDrive site; pass `site` (a site URL) to use another.

## Setup

1. **Create an Entra app registration** (Entra admin center → App registrations → New registration).
   - Supported account types: your organization only.
   - Add a **Mobile and desktop** redirect URI: `http://localhost`.
   - Add the delegated Microsoft Graph permissions listed below.
2. **Install and configure** (Node 20.12+):
   ```powershell
   npm install
   copy .env.example .env     # then fill in your client ID, tenant ID and SharePoint host
   npm run login              # one-time browser sign-in
   ```
   The token cache is stored at `~/.graph-mcp/token_cache.bin` (DPAPI-encrypted on Windows).
3. **Register with your MCP client**, e.g. Claude Code:
   ```powershell
   claude mcp add graph --scope user -- node C:\path\to\graph-mcp\src\server.js
   ```

## Permissions

Delegated Microsoft Graph permissions, by area:

| Area | Permission |
|---|---|
| OneNote | `Notes.ReadWrite.All` |
| Mail drafts | `Mail.ReadWrite` |
| Teams chats | `Chat.Read` |
| Lists | `Sites.ReadWrite.All` |

Optional features are off by default. Add the permission in your app registration **first**, then flip the switch in `src/auth.js` and re-run `npm run login`. Requesting a permission that isn't granted breaks sign-in for every tool.

| Switch (`src/auth.js`) | Enables | Extra permission |
|---|---|---|
| `TEAMS_CHANNELS` | Teams channel tools | Graph: `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All` (admin consent) |
| `LISTS_MANAGE` | `lists_create_list`, `lists_add_column`, `lists_update_column` | Graph: `Sites.Manage.All` |
| `LISTS_VIEWS` | `lists_get_views`, `lists_create_view`, and hyperlink columns/values | SharePoint (Office 365 SharePoint Online): `AllSites.Manage` |

Microsoft Graph has no API for list views and can't create, label or write URL (hyperlink) columns, so those features call SharePoint REST with a separate token.

## Configuration

Read from `.env` (git-ignored; see `.env.example`) or from real environment variables, which win:

| Variable | Notes |
|---|---|
| `GRAPH_CLIENT_ID`, `GRAPH_TENANT_ID` | Required |
| `SHAREPOINT_ROOT_HOST` | Required when `LISTS_VIEWS` is on, e.g. `contoso.sharepoint.com` |
| `GRAPH_SCOPES`, `GRAPH_TOKEN_CACHE` | Optional overrides |

## Adding a Graph feature

1. Add the delegated permission to your app registration.
2. Add its scope to `DEFAULT_SCOPES` in `src/auth.js`, then re-run `npm run login` to consent.
3. Create `src/tools/<area>.js` exporting `register(server)` and add it to `MODULES` in `src/server.js`.
