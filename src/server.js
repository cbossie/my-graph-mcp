#!/usr/bin/env node
// Graph MCP server (stdio). Add new areas by creating tools/<area>.js exporting register(server) and listing it below.
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import * as onenote from "./tools/onenote.js";
import * as mail from "./tools/mail.js";
import * as teams from "./tools/teams.js";
import * as lists from "./tools/lists.js";

const MODULES = [onenote, mail, teams, lists]; // e.g. add: calendar, mail, planner... (and add their scopes in auth.js)

const server = new McpServer({ name: "graph", version: "0.1.0" });
for (const m of MODULES) m.register(server);

await server.connect(new StdioServerTransport());
