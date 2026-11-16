#!/usr/bin/env node
// Graph MCP server (stdio). Add new areas by creating tools/<area>.js exporting register(server) and listing it below.
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import * as onenote from "./tools/onenote.js";
import * as mail from "./tools/mail.js";
import * as teams from "./tools/teams.js";
import * as lists from "./tools/lists.js";
import * as calendar from "./tools/calendar.js";
import * as planner from "./tools/planner.js";
import { CALENDAR, PLANNER } from "./auth.js";

const MODULES = [onenote, mail, teams, lists, ...(CALENDAR ? [calendar] : []), ...(PLANNER ? [planner] : [])]; // add new areas here (and their scopes in auth.js)

const server = new McpServer({ name: "graph", version: "0.1.0" });
for (const m of MODULES) m.register(server);

await server.connect(new StdioServerTransport());
