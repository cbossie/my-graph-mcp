// OneNote tools. Read + create/append only; deliberately no delete or overwrite.
import { z } from "zod";
import { request } from "../graph.js";

const escapeHtml = (s) =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

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

const ParentType = z.enum(["notebook", "sectionGroup"]);
const parentPath = (type, id) =>
  `me/onenote/${type === "notebook" ? "notebooks" : "sectionGroups"}/${id}`;

const json = (value) => ({ content: [{ type: "text", text: JSON.stringify(value, null, 2) }] });
const text = (value) => ({ content: [{ type: "text", text: value }] });

export function register(server) {
  server.tool("onenote_list_notebooks", "List OneNote notebooks (id, name).", {}, async () => {
    const r = await request("GET", "me/onenote/notebooks", { params: { $select: "id,displayName" } });
    return json(r.value.map((n) => ({ id: n.id, name: n.displayName })));
  });

  server.tool(
    "onenote_list_section_groups",
    "List section groups directly inside a notebook or another section group.",
    { parent_id: z.string(), parent_type: ParentType.default("notebook") },
    async ({ parent_id, parent_type }) => {
      const r = await request("GET", `${parentPath(parent_type, parent_id)}/sectionGroups`, {
        params: { $select: "id,displayName" },
      });
      return json(r.value.map((g) => ({ id: g.id, name: g.displayName })));
    },
  );

  server.tool(
    "onenote_list_sections",
    "List sections directly inside a notebook or section group (not recursive; use onenote_list_section_groups to descend).",
    { parent_id: z.string(), parent_type: ParentType.default("notebook") },
    async ({ parent_id, parent_type }) => {
      const r = await request("GET", `${parentPath(parent_type, parent_id)}/sections`, {
        params: { $select: "id,displayName" },
      });
      return json(r.value.map((s) => ({ id: s.id, name: s.displayName })));
    },
  );

  server.tool(
    "onenote_find_by_path",
    "Resolve a slash-separated path (e.g. 'Projects/Active/Website Redesign') inside a notebook to a section group or section. Each segment is matched case-insensitively against section groups first, then sections.",
    { notebook_id: z.string(), path: z.string() },
    async ({ notebook_id, path }) => {
      let current = { id: notebook_id, type: "notebook", name: "(notebook)" };
      for (const segment of path.split("/").map((s) => s.trim()).filter(Boolean)) {
        const base = parentPath(current.type, current.id);
        const [groups, sections] = await Promise.all([
          request("GET", `${base}/sectionGroups`, { params: { $select: "id,displayName" } }),
          request("GET", `${base}/sections`, { params: { $select: "id,displayName" } }),
        ]);
        const match = (x) => x.displayName.toLowerCase() === segment.toLowerCase();
        const g = groups.value.find(match);
        const s = g ? undefined : sections.value.find(match);
        if (!g && !s) {
          const available = [...groups.value, ...sections.value].map((x) => x.displayName);
          throw new Error(`'${segment}' not found under '${current.name}'. Available: ${available.join(", ")}`);
        }
        current = g
          ? { id: g.id, type: "sectionGroup", name: g.displayName }
          : { id: s.id, type: "section", name: s.displayName };
      }
      return json(current);
    },
  );

  server.tool(
    "onenote_create_section",
    "Create a new section inside a notebook or section group. Fails if a section with that name already exists there.",
    { parent_id: z.string(), parent_type: ParentType.default("notebook"), name: z.string() },
    async ({ parent_id, parent_type, name }) => {
      const r = await request("POST", `${parentPath(parent_type, parent_id)}/sections`, {
        json: { displayName: name },
      });
      return json({ id: r.id, name: r.displayName });
    },
  );

  server.tool(
    "onenote_list_pages",
    "List pages in a section, newest first (id, title, lastModified).",
    { section_id: z.string(), top: z.number().int().min(1).max(100).default(50) },
    async ({ section_id, top }) => {
      const r = await request("GET", `me/onenote/sections/${section_id}/pages`, {
        params: { $select: "id,title,lastModifiedDateTime", $orderby: "lastModifiedDateTime desc", $top: top },
      });
      return json(r.value.map((p) => ({ id: p.id, title: p.title, lastModified: p.lastModifiedDateTime })));
    },
  );

  server.tool(
    "onenote_get_page",
    "Get a page's content. as_text=true strips HTML tags; false returns raw HTML.",
    { page_id: z.string(), as_text: z.boolean().default(true) },
    async ({ page_id, as_text }) => {
      const body = await request("GET", `me/onenote/pages/${page_id}/content`);
      return text(as_text ? htmlToText(body) : body);
    },
  );

  server.tool(
    "onenote_create_page",
    "Create a new page in a section. body_html is HTML for the page body (e.g. <p>text</p>).",
    { section_id: z.string(), title: z.string(), body_html: z.string() },
    async ({ section_id, title, body_html }) => {
      const doc = `<!DOCTYPE html><html><head><title>${escapeHtml(title)}</title></head><body>${body_html}</body></html>`;
      const r = await request("POST", `me/onenote/sections/${section_id}/pages`, {
        body: doc,
        headers: { "Content-Type": "text/html" },
      });
      return json({ id: r.id, title: r.title, url: r.links?.oneNoteWebUrl?.href });
    },
  );

  server.tool(
    "onenote_append_to_page",
    "Append HTML to the end of an existing page. Existing content is not changed.",
    { page_id: z.string(), body_html: z.string() },
    async ({ page_id, body_html }) => {
      await request("PATCH", `me/onenote/pages/${page_id}/content`, {
        json: [{ target: "body", action: "append", content: body_html }],
      });
      return text("Appended.");
    },
  );
}
