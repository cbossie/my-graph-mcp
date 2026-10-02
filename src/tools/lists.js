// Microsoft Lists (SharePoint lists) tools.
// Items/columns/getList use Graph (Sites.ReadWrite.All). createList needs Sites.Manage.All (LISTS_MANAGE).
// Views are not in Graph, so getViews/createView use SharePoint REST (LISTS_VIEWS).
import { z } from "zod";
import { LISTS_MANAGE, LISTS_VIEWS } from "../auth.js";
import { request, sharePointRequest } from "../graph.js";

const json = (value) => ({ content: [{ type: "text", text: JSON.stringify(value, null, 2) }] });

const SiteParam = z
  .string()
  .optional()
  .describe(
    "Site to use: omitted = your personal site (where lists made in the Lists app live), 'root', a site URL like https://tenant.sharepoint.com/sites/Team, or a Graph site id.",
  );
const ListParam = z.string().describe("List id (GUID) or list display name.");
const FieldsParam = z
  .record(z.any())
  .describe(
    "Column values keyed by column INTERNAL name (see lists_get_columns), e.g. { Title: 'Call Matt', Status: 'Active' }. Hyperlink columns take { Url, Description }. Date-only columns take 'YYYY-MM-DD'.",
  );

const siteCache = new Map();

/** Resolve the site param to { id, webUrl }. */
async function resolveSite(site) {
  const key = site ?? "";
  if (siteCache.has(key)) return siteCache.get(key);

  let s;
  if (!site) {
    const drive = await request("GET", "me/drive", { params: { $select: "webUrl" } });
    const u = new URL(drive.webUrl);
    const [kind, name] = u.pathname.split("/").filter(Boolean); // /personal/<user>/Documents
    s = await request("GET", `sites/${u.host}:/${kind}/${name}`, { params: { $select: "id,webUrl" } });
  } else if (site === "root") {
    s = await request("GET", "sites/root", { params: { $select: "id,webUrl" } });
  } else if (/^https?:\/\//i.test(site)) {
    const u = new URL(site);
    const p = u.pathname.replace(/\/$/, "");
    s = await request("GET", p ? `sites/${u.host}:${p}` : `sites/${u.host}`, { params: { $select: "id,webUrl" } });
  } else {
    s = await request("GET", `sites/${site}`, { params: { $select: "id,webUrl" } });
  }
  const resolved = { id: s.id, webUrl: s.webUrl };
  siteCache.set(key, resolved);
  return resolved;
}

const listPath = async (site, list) => `sites/${(await resolveSite(site)).id}/lists/${encodeURIComponent(list)}`;

const cleanFields = (fields = {}) =>
  Object.fromEntries(Object.entries(fields).filter(([k]) => !k.startsWith("@odata")));

const itemSummary = (i) => ({
  id: i.id,
  created: i.createdDateTime,
  modified: i.lastModifiedDateTime,
  webUrl: i.webUrl,
  fields: cleanFields(i.fields),
});

const COLUMN_TYPES = ["text", "note", "number", "boolean", "dateTime", "currency", "choice", "hyperlink"];
const GRAPH_ONLY_TYPES = ["lookup", "personOrGroup", "calculated", "term"];

// Graph has no "note" or "hyperlink" type: multi-line text is a text column with allowMultipleLines,
// and links are hyperlinkOrPicture with isPicture=false.
const columnType = (c) => {
  if (c.text?.allowMultipleLines) return "note";
  if (c.hyperlinkOrPicture) return c.hyperlinkOrPicture.isPicture ? "picture" : "hyperlink";
  return COLUMN_TYPES.concat(GRAPH_ONLY_TYPES).find((t) => c[t]);
};

const ColumnSchema = z.object({
  name: z.string().describe("Column name (becomes the internal name)."),
  type: z.enum(COLUMN_TYPES),
  choices: z.array(z.string()).optional().describe("Required when type is 'choice'."),
  required: z.boolean().default(false),
  date_only: z.boolean().default(false).describe("Only for type 'dateTime': date without a time."),
  description: z.string().optional(),
  default_value: z.string().optional().describe("Default value as text (e.g. a choice label)."),
  display_as: z
    .enum(["dropDownMenu", "radioButtons", "checkBoxes"])
    .optional()
    .describe("Only for type 'choice' (default dropDownMenu)."),
});

const columnTypeDefinition = (c) => {
  switch (c.type) {
    case "note":
      return { text: { allowMultipleLines: true, textType: "plain" } };
    case "choice":
      return { choice: { choices: c.choices ?? [], allowTextEntry: false, displayAs: c.display_as ?? "dropDownMenu" } };
    case "hyperlink":
      return { hyperlinkOrPicture: { isPicture: false } };
    case "dateTime":
      return { dateTime: c.date_only ? { format: "dateOnly", displayAs: "default" } : {} };
    default:
      return { [c.type]: {} };
  }
};

const columnDefinition = (c) => ({
  name: c.name,
  required: c.required,
  ...(c.description !== undefined && { description: c.description }),
  ...(c.default_value !== undefined && { defaultValue: { value: c.default_value } }),
  ...columnTypeDefinition(c),
});

const listGuid = async (site, list) =>
  /^[0-9a-f-]{36}$/i.test(list)
    ? list
    : (await request("GET", await listPath(site, list), { params: { $select: "id" } })).id;

// Graph can't create or read URL columns properly (it makes/returns a typeless column), so hyperlink
// columns go through SharePoint REST, which needs the AllSites.Manage permission (LISTS_VIEWS).
const requireRest = () => {
  if (!LISTS_VIEWS) {
    throw new Error(
      "Hyperlink columns use SharePoint REST: set LISTS_VIEWS = true in src/auth.js (needs AllSites.Manage) and re-run `npm run login`.",
    );
  }
};

// Graph can't write URL field values either (400/500), so { Url, Description } values are split out of
// `fields` and written with SharePoint REST validateUpdateListItem, which understands URL fields.
const isUrlValue = (v) => v && typeof v === "object" && typeof v.Url === "string";

const splitUrlFields = (fields) => {
  const urlFields = {};
  const rest = {};
  for (const [k, v] of Object.entries(fields)) (isUrlValue(v) ? urlFields : rest)[k] = v;
  return { urlFields, rest };
};

async function applyUrlFields(site, list, itemId, urlFields) {
  requireRest();
  const { webUrl } = await resolveSite(site);
  const id = await listGuid(site, list);
  const r = await sharePointRequest(webUrl, "POST", `web/lists/getbyid('${id}')/items(${itemId})/validateupdatelistitem`, {
    json: {
      formValues: Object.entries(urlFields).map(([FieldName, v]) => ({
        FieldName,
        FieldValue: v.Description ? `${v.Url}, ${v.Description}` : v.Url,
      })),
      bNewDocumentUpdate: false,
    },
  });
  const results = r?.ValidateUpdateListItem?.results ?? r?.ValidateUpdateListItem ?? [];
  const failed = results.filter((x) => x.HasException);
  if (failed.length) throw new Error(`Could not set ${failed.map((x) => `${x.FieldName}: ${x.ErrorMessage}`).join("; ")}`);
}

async function addHyperlinkColumn(site, list, c) {
  requireRest();
  const { webUrl } = await resolveSite(site);
  const id = await listGuid(site, list);
  const f = await sharePointRequest(webUrl, "POST", `web/lists/getbyid('${id}')/fields`, {
    json: {
      __metadata: { type: "SP.FieldUrl" },
      Title: c.name,
      FieldTypeKind: 11,
      DisplayFormat: 0, // 0 = hyperlink, 1 = picture
      Required: c.required,
      ...(c.description !== undefined && { Description: c.description }),
      ...(c.default_value !== undefined && { DefaultValue: c.default_value }),
    },
  });
  return { name: f.InternalName, displayName: f.Title, type: "hyperlink" };
}

export function register(server) {
  if (LISTS_MANAGE) {
    server.tool(
      "lists_create_list",
      "Create a new list on a site, optionally with columns. (Needs Sites.Manage.All.)",
      {
        display_name: z.string(),
        description: z.string().optional(),
        template: z.enum(["genericList", "documentLibrary"]).default("genericList"),
        columns: z.array(ColumnSchema).default([]),
        site: SiteParam,
      },
      async ({ display_name, description, template, columns, site }) => {
        const links = columns.filter((c) => c.type === "hyperlink");
        if (links.length) requireRest(); // fail before creating anything
        const body = {
          displayName: display_name,
          description,
          list: { template },
          columns: columns.filter((c) => c.type !== "hyperlink").map(columnDefinition),
        };
        const r = await request("POST", `sites/${(await resolveSite(site)).id}/lists`, { json: body });
        for (const c of links) await addHyperlinkColumn(site, r.id, c);
        return json({ id: r.id, name: r.displayName, webUrl: r.webUrl });
      },
    );

    server.tool(
      "lists_add_column",
      "Add a column to an existing list. Uses the same column schema as lists_create_list. (Needs Sites.Manage.All.)",
      { list: ListParam, column: ColumnSchema, site: SiteParam },
      async ({ list, column, site }) => {
        if (column.type === "hyperlink") return json(await addHyperlinkColumn(site, list, column));
        const r = await request("POST", `${await listPath(site, list)}/columns`, { json: columnDefinition(column) });
        return json({ name: r.name, displayName: r.displayName, type: columnType(r) });
      },
    );

    server.tool(
      "lists_update_column",
      "Change an existing column: choices, required, date-only vs date+time, description, default value, or choice display style. `column` is the column's internal name, display name or id. Only the given properties change. (Needs Sites.Manage.All.)",
      {
        list: ListParam,
        column: z.string(),
        choices: z.array(z.string()).optional().describe("Replaces the full choice list (choice columns only)."),
        required: z.boolean().optional(),
        date_only: z.boolean().optional().describe("dateTime columns only."),
        description: z.string().optional(),
        default_value: z.string().optional(),
        display_as: z.enum(["dropDownMenu", "radioButtons", "checkBoxes"]).optional().describe("Choice columns only."),
        site: SiteParam,
      },
      async ({ list, column, choices, required, date_only, description, default_value, display_as, site }) => {
        const base = `${await listPath(site, list)}/columns`;
        const all = await request("GET", base);
        const existing = all.value.find((c) => c.id === column || c.name === column || c.displayName === column);
        if (!existing) throw new Error(`Column '${column}' not found. Available: ${all.value.filter((c) => !c.readOnly).map((c) => c.name).join(", ")}`);

        const patch = {};
        if (required !== undefined) patch.required = required;
        if (description !== undefined) patch.description = description;
        if (default_value !== undefined) patch.defaultValue = { value: default_value };
        if (choices !== undefined || display_as !== undefined) {
          if (!existing.choice) throw new Error(`'${column}' is not a choice column.`);
          patch.choice = {
            ...existing.choice,
            ...(choices !== undefined && { choices }),
            ...(display_as !== undefined && { displayAs: display_as }),
          };
        }
        if (date_only !== undefined) {
          if (!existing.dateTime) throw new Error(`'${column}' is not a dateTime column.`);
          patch.dateTime = { ...existing.dateTime, format: date_only ? "dateOnly" : "dateTime" };
        }
        const r = await request("PATCH", `${base}/${existing.id}`, { json: patch });
        return json({ name: r.name, displayName: r.displayName, type: columnType(r), required: r.required });
      },
    );
  }

  server.tool(
    "lists_get_list",
    "Get one list's details. If `list` is omitted, returns all visible lists on the site (use this to discover lists).",
    { list: ListParam.optional(), site: SiteParam },
    async ({ list, site }) => {
      const base = `sites/${(await resolveSite(site)).id}/lists`;
      if (!list) {
        const r = await request("GET", base, { params: { $select: "id,displayName,webUrl,list" } });
        return json(
          r.value
            .filter((l) => !l.list?.hidden)
            .map((l) => ({ id: l.id, name: l.displayName, template: l.list?.template, webUrl: l.webUrl })),
        );
      }
      const l = await request("GET", `${base}/${encodeURIComponent(list)}`, {
        params: { $select: "id,displayName,description,webUrl,createdDateTime,lastModifiedDateTime,list" },
      });
      return json({
        id: l.id,
        name: l.displayName,
        description: l.description,
        template: l.list?.template,
        webUrl: l.webUrl,
        created: l.createdDateTime,
        modified: l.lastModifiedDateTime,
      });
    },
  );

  server.tool(
    "lists_get_columns",
    "Get a list's columns: internal name (use this in fields/filters), display name, type, required, choices. Hides read-only/system columns unless include_hidden=true.",
    { list: ListParam, include_hidden: z.boolean().default(false), site: SiteParam },
    async ({ list, include_hidden, site }) => {
      const r = await request("GET", `${await listPath(site, list)}/columns`);
      const visible = r.value.filter((c) => include_hidden || (!c.hidden && !c.readOnly));

      // Graph returns URL columns with no type facet; when SharePoint REST is available, use it to label them.
      let restTypes = {};
      if (LISTS_VIEWS && visible.some((c) => !columnType(c))) {
        const { webUrl } = await resolveSite(site);
        const id = await listGuid(site, list);
        const fields = await sharePointRequest(webUrl, "GET", `web/lists/getbyid('${id}')/fields?$select=InternalName,TypeAsString`);
        restTypes = Object.fromEntries(fields.map((f) => [f.InternalName, f.TypeAsString === "URL" ? "hyperlink" : f.TypeAsString]));
      }

      return json(
        visible
          .map((c) => ({
            name: c.name,
            displayName: c.displayName,
            type: columnType(c) ?? restTypes[c.name],
            required: c.required,
            readOnly: c.readOnly,
            choices: c.choice?.choices,
            displayAs: c.choice?.displayAs,
            format: c.dateTime?.format,
            description: c.description || undefined,
            defaultValue: c.defaultValue?.value,
            id: c.id,
          })),
      );
    },
  );

  server.tool(
    "lists_create_item",
    "Create an item in a list.",
    { list: ListParam, fields: FieldsParam, site: SiteParam },
    async ({ list, fields, site }) => {
      const { urlFields, rest } = splitUrlFields(fields);
      const path = await listPath(site, list);
      let r = await request("POST", `${path}/items`, { json: { fields: rest } });
      if (Object.keys(urlFields).length) {
        await applyUrlFields(site, list, r.id, urlFields);
        r = await request("GET", `${path}/items/${r.id}`, { params: { $expand: "fields" } });
      }
      return json(itemSummary(r));
    },
  );

  server.tool(
    "lists_get_item",
    "Get one list item by id.",
    { list: ListParam, item_id: z.string(), site: SiteParam },
    async ({ list, item_id, site }) => {
      const r = await request("GET", `${await listPath(site, list)}/items/${item_id}`, {
        params: { $expand: "fields" },
      });
      return json(itemSummary(r));
    },
  );

  server.tool(
    "lists_update_item",
    "Update columns on a list item. Only the given fields change.",
    { list: ListParam, item_id: z.string(), fields: FieldsParam, site: SiteParam },
    async ({ list, item_id, fields, site }) => {
      const { urlFields, rest } = splitUrlFields(fields);
      const path = await listPath(site, list);
      let r;
      if (Object.keys(rest).length) r = await request("PATCH", `${path}/items/${item_id}/fields`, { json: rest });
      if (Object.keys(urlFields).length) {
        await applyUrlFields(site, list, item_id, urlFields);
        r = (await request("GET", `${path}/items/${item_id}`, { params: { $expand: "fields" } })).fields;
      }
      return json({ id: item_id, fields: cleanFields(r) });
    },
  );

  server.tool(
    "lists_delete_item",
    "Delete a list item (goes to the site recycle bin). Show the user which item will be deleted and get their OK before calling with confirm=true.",
    { list: ListParam, item_id: z.string(), confirm: z.literal(true), site: SiteParam },
    async ({ list, item_id, site }) => {
      await request("DELETE", `${await listPath(site, list)}/items/${item_id}`);
      return json({ deleted: item_id });
    },
  );

  server.tool(
    "lists_query_items",
    "Query list items. filter/order_by use OData on fields, e.g. filter=\"fields/Status eq 'Active'\", order_by=\"fields/Modified desc\". Returns nextLink when there are more results; pass it back as next_link.",
    {
      list: ListParam,
      filter: z.string().optional(),
      order_by: z.string().optional(),
      select: z.array(z.string()).optional().describe("Column internal names to return (default: all)."),
      top: z.number().int().min(1).max(200).default(50),
      next_link: z.string().optional(),
      site: SiteParam,
    },
    async ({ list, filter, order_by, select, top, next_link, site }) => {
      let r;
      if (next_link) {
        if (!next_link.startsWith("https://graph.microsoft.com/")) throw new Error("next_link must be a Graph URL");
        r = await request("GET", next_link, { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } });
      } else {
        const params = {
          $expand: select?.length ? `fields($select=${select.join(",")})` : "fields",
          $top: top,
        };
        if (filter) params.$filter = filter;
        if (order_by) params.$orderby = order_by;
        r = await request("GET", `${await listPath(site, list)}/items`, {
          params,
          // Lets Graph filter/sort on columns that aren't indexed (fine for small personal lists).
          headers: filter || order_by ? { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } : undefined,
        });
      }
      return json({ items: r.value.map(itemSummary), nextLink: r["@odata.nextLink"] });
    },
  );

  if (!LISTS_VIEWS) return;

  server.tool(
    "lists_get_views",
    "Get a list's views (title, default, row limit, CAML query, fields). Uses SharePoint REST.",
    { list: ListParam, site: SiteParam },
    async ({ list, site }) => {
      const { webUrl } = await resolveSite(site);
      const id = await listGuid(site, list);
      const views = await sharePointRequest(
        webUrl,
        "GET",
        `web/lists/getbyid('${id}')/views?$select=Id,Title,DefaultView,Hidden,PersonalView,RowLimit,ViewQuery&$expand=ViewFields`,
      );
      return json(
        views.map((v) => ({
          id: v.Id,
          title: v.Title,
          isDefault: v.DefaultView,
          hidden: v.Hidden,
          personal: v.PersonalView,
          rowLimit: v.RowLimit,
          query: v.ViewQuery,
          fields: v.ViewFields?.Items?.results ?? v.ViewFields?.Items,
        })),
      );
    },
  );

  server.tool(
    "lists_create_view",
    "Create a view on a list. `query` is CAML (the <Where>/<OrderBy> part), e.g. \"<Where><Eq><FieldRef Name='Status'/><Value Type='Choice'>Active</Value></Eq></Where>\". Uses SharePoint REST.",
    {
      list: ListParam,
      title: z.string(),
      fields: z.array(z.string()).default([]).describe("Column internal names to show, in order (default: SharePoint's defaults)."),
      query: z.string().default(""),
      row_limit: z.number().int().min(1).max(500).default(30),
      make_default: z.boolean().default(false),
      personal: z.boolean().default(false),
      site: SiteParam,
    },
    async ({ list, title, fields, query, row_limit, make_default, personal, site }) => {
      const { webUrl } = await resolveSite(site);
      const id = await listGuid(site, list);
      const base = `web/lists/getbyid('${id}')/views`;
      const view = await sharePointRequest(webUrl, "POST", base, {
        json: {
          __metadata: { type: "SP.View" },
          Title: title,
          ViewQuery: query,
          RowLimit: row_limit,
          DefaultView: make_default,
          ...(personal && { PersonalView: true }),
        },
      });
      if (fields.length) {
        const viewFields = `${base}/getbyid('${view.Id}')/viewfields`;
        await sharePointRequest(webUrl, "POST", `${viewFields}/removeallviewfields`);
        for (const f of fields) {
          await sharePointRequest(webUrl, "POST", `${viewFields}/addviewfield('${f.replace(/'/g, "''")}')`);
        }
      }
      return json({ id: view.Id, title: view.Title });
    },
  );
}
