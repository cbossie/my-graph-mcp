// Teams tools. Read-only: deliberately no send tools (no Send scopes are requested either).
import { z } from "zod";
import { TEAMS_CHANNELS } from "../auth.js";
import { request } from "../graph.js";

const unescapeHtml = (s) =>
  s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&");

const htmlToText = (html) =>
  unescapeHtml(
    html
      .replace(/<(script|style)[\s\S]*?<\/\1>/gi, "")
      .replace(/<br\s*\/?>|<\/(p|div|li|h\d)>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/\n{3,}/g, "\n\n"),
  ).trim();

const json = (value) => ({ content: [{ type: "text", text: JSON.stringify(value, null, 2) }] });

const messageSummary = (m) => ({
  id: m.id,
  from: m.from?.user?.displayName ?? m.from?.application?.displayName ?? "(system)",
  created: m.createdDateTime,
  text: m.body?.contentType === "html" ? htmlToText(m.body.content ?? "") : (m.body?.content ?? ""),
  attachments: (m.attachments ?? []).map((a) => a.name).filter(Boolean),
});

const visible = (r) => r.value.filter((m) => m.messageType === "message" && !m.deletedDateTime);

export function register(server) {
  server.tool(
    "teams_list_chats",
    "List your Teams chats (1:1, group, meeting), most recently active first, with members and a preview of the last message.",
    { top: z.number().int().min(1).max(50).default(25) },
    async ({ top }) => {
      const r = await request("GET", "me/chats", {
        params: { $expand: "members,lastMessagePreview", $orderby: "lastMessagePreview/createdDateTime desc", $top: top },
      });
      return json(
        r.value.map((c) => ({
          id: c.id,
          type: c.chatType,
          topic: c.topic,
          members: (c.members ?? []).map((m) => m.displayName),
          lastMessage: c.lastMessagePreview && {
            created: c.lastMessagePreview.createdDateTime,
            from: c.lastMessagePreview.from?.user?.displayName,
            text:
              c.lastMessagePreview.body?.contentType === "html"
                ? htmlToText(c.lastMessagePreview.body.content ?? "")
                : c.lastMessagePreview.body?.content,
          },
        })),
      );
    },
  );

  server.tool(
    "teams_get_chat_messages",
    "Read recent messages in a chat, newest first.",
    { chat_id: z.string(), top: z.number().int().min(1).max(50).default(20) },
    async ({ chat_id, top }) => {
      const r = await request("GET", `me/chats/${chat_id}/messages`, { params: { $top: top } });
      return json(visible(r).map(messageSummary));
    },
  );

  if (!TEAMS_CHANNELS) return;

  server.tool("teams_list_teams", "List the teams you belong to.", {}, async () => {
    const r = await request("GET", "me/joinedTeams", { params: { $select: "id,displayName" } });
    return json(r.value.map((t) => ({ id: t.id, name: t.displayName })));
  });

  server.tool(
    "teams_list_channels",
    "List channels in a team.",
    { team_id: z.string() },
    async ({ team_id }) => {
      const r = await request("GET", `teams/${team_id}/channels`, { params: { $select: "id,displayName" } });
      return json(r.value.map((c) => ({ id: c.id, name: c.displayName })));
    },
  );

  server.tool(
    "teams_get_channel_messages",
    "Read recent top-level messages in a channel, newest first.",
    { team_id: z.string(), channel_id: z.string(), top: z.number().int().min(1).max(50).default(20) },
    async ({ team_id, channel_id, top }) => {
      const r = await request("GET", `teams/${team_id}/channels/${channel_id}/messages`, { params: { $top: top } });
      return json(visible(r).map(messageSummary));
    },
  );
}
