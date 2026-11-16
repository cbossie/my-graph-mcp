// Microsoft Planner tools (Tasks.ReadWrite + Group.Read.All; PLANNER switch in auth.js).
//
// Deliberately NOT here: deleting tasks, buckets or plans.
//
// Things Planner does differently from the rest of Graph:
//   - Optimistic concurrency: PATCH needs If-Match with the entity's current @odata.etag. The task and
//     its details (description, checklist, references) are separate resources with separate etags.
//     The tools fetch the etag themselves and retry once on a conflict. Planner reports a stale etag
//     as 412 or 409 depending on the endpoint, so both count.
//   - Dates are UTC dateTimes. Tools take YYYY-MM-DD and store 12:00 UTC that day, which shows as the
//     same calendar date for anyone between UTC-11 and UTC+11 (midnight local would show as the day
//     before for viewers west of you). Dates are read back as the date in PLANNER_TZ.
//   - percentComplete is 0 / 50 / 100 in the UI; priority 0-10 maps to urgent/important/medium/low.
import { randomUUID } from "node:crypto";
import { z } from "zod";
import { request } from "../graph.js";

const json = (value) => ({ content: [{ type: "text", text: JSON.stringify(value, null, 2) }] });

export const PLANNER_TZ =
  process.env.PLANNER_TIME_ZONE ||
  process.env.CALENDAR_TIME_ZONE ||
  Intl.DateTimeFormat().resolvedOptions().timeZone ||
  "UTC";

const enc = encodeURIComponent;
const isGuid = (s) => /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s);

// ---------- IDs from URLs ----------

/** Plan id from a bare id or a Planner URL (.../plan/<id>/..., or ?planId=<id> on older tasks.office.com links). */
export function parsePlanId(input) {
  const s = input.trim();
  if (!/^https?:\/\//i.test(s) && !s.includes("/")) return s;
  const u = new URL(s.replace("#/", "")); // old links keep the route in the hash
  const fromParam = u.searchParams.get("planId");
  if (fromParam) return fromParam;
  const m = (u.pathname + u.hash).match(/\/plan\/([^/?#]+)/i);
  if (m) return decodeURIComponent(m[1]);
  throw new Error(`Could not find a plan id in "${input}". Pass the id (the segment after /plan/ in the URL).`);
}

/** Task id from a bare id or a Planner task URL (.../task/<id>, or ?taskId=<id>). */
export function parseTaskId(input) {
  const s = input.trim();
  if (!/^https?:\/\//i.test(s) && !s.includes("/")) return s;
  const u = new URL(s.replace("#/", ""));
  const fromParam = u.searchParams.get("taskId");
  if (fromParam) return fromParam;
  const m = (u.pathname + u.hash).match(/\/task\/([^/?#]+)/i);
  if (m) return decodeURIComponent(m[1]);
  throw new Error(`Could not find a task id in "${input}".`);
}

// ---------- dates (YYYY-MM-DD -> 12:00 UTC; UTC -> date in PLANNER_TZ) ----------

const YMD = /^\d{4}-\d{2}-\d{2}$/;

/** '2026-11-23' -> '2026-11-23T12:00:00Z'. */
export function dateToUtc(ymd) {
  if (!YMD.test(ymd) || new Date(`${ymd}T12:00:00Z`).toISOString().slice(0, 10) !== ymd) throw new Error(`Dates must be YYYY-MM-DD, got "${ymd}".`);
  return `${ymd}T12:00:00Z`;
}

/** UTC dateTime -> 'YYYY-MM-DD' as seen in PLANNER_TZ. */
export function utcToDate(iso) {
  if (!iso) return undefined;
  return new Intl.DateTimeFormat("en-CA", { timeZone: PLANNER_TZ, year: "numeric", month: "2-digit", day: "2-digit" }).format(
    new Date(iso),
  );
}

// ---------- priority / progress ----------

const PRIORITY_VALUE = { urgent: 1, important: 3, medium: 5, low: 9 };
export const priorityName = (n) => (n == null ? undefined : n <= 1 ? "urgent" : n <= 4 ? "important" : n <= 7 ? "medium" : "low");
const priorityOut = (n) => (n == null ? undefined : { name: priorityName(n), value: n });

const PriorityParam = z
  .union([z.enum(["urgent", "important", "medium", "low"]), z.number().int().min(0).max(10)])
  .describe("urgent | important | medium | low (stored as 1, 3, 5, 9), or a raw 0-10 value (0-1 urgent, 2-4 important, 5-7 medium, 8-10 low).");
const priorityIn = (p) => (typeof p === "number" ? p : PRIORITY_VALUE[p]);

const PercentParam = z
  .union([z.enum(["not_started", "in_progress", "completed"]), z.number().int().min(0).max(100)])
  .describe("not_started | in_progress | completed, or 0 / 50 / 100. Planner's UI only has those three; other numbers are normalised (1-99 -> 50).");
const percentIn = (p) => {
  if (p === "not_started") return 0;
  if (p === "in_progress") return 50;
  if (p === "completed") return 100;
  return p === 0 || p === 100 ? p : 50;
};
const statusOf = (pc) => (pc >= 100 ? "completed" : pc > 0 ? "in_progress" : "not_started");

// ---------- Graph helpers ----------

async function getAll(path, params) {
  const out = [];
  let r = await request("GET", path, { params });
  out.push(...r.value);
  while (r["@odata.nextLink"]) {
    r = await request("GET", r["@odata.nextLink"]);
    out.push(...r.value);
  }
  return out;
}

const REPRESENTATION = { Prefer: "return=representation" };

/**
 * PATCH a Planner entity with optimistic concurrency. `build(current)` returns the patch body (or null
 * for nothing to do) from the freshly fetched entity, so a retry after a 412 rebuilds from new state.
 */
async function patchWithEtag(path, build, label) {
  for (let attempt = 1; ; attempt++) {
    const current = await request("GET", path);
    const body = build(current);
    if (!body || !Object.keys(body).length) return current;
    try {
      return await request("PATCH", path, {
        json: body,
        headers: { "If-Match": current["@odata.etag"], ...REPRESENTATION },
      });
    } catch (e) {
      const conflict = e.status === 412 || e.status === 409;
      if (!conflict || attempt >= 2) {
        if (conflict) {
          e.message = `Conflict updating ${label}: it changed again while saving, even after re-reading it once (HTTP ${e.status}, etag mismatch). Nothing from this call was saved to the ${label}; read it again and retry. ${e.message}`;
        }
        throw e;
      }
      // someone else changed it between our GET and PATCH: re-read and try once more
    }
  }
}

/** Task details can 404 for a moment right after the task is created. */
async function waitForDetails(taskId) {
  for (let i = 0; ; i++) {
    try {
      return await request("GET", `planner/tasks/${enc(taskId)}/details`);
    } catch (e) {
      if (e.status !== 404 || i >= 4) throw e;
      await new Promise((r) => setTimeout(r, 500 * (i + 1)));
    }
  }
}

// ---------- plan context (buckets, category labels, group members), cached per call ----------

const groupNameCache = new Map();
async function groupName(groupId) {
  if (!groupId) return undefined;
  if (!groupNameCache.has(groupId)) {
    groupNameCache.set(
      groupId,
      request("GET", `groups/${enc(groupId)}`, { params: { $select: "displayName" } })
        .then((g) => g.displayName)
        .catch(() => undefined),
    );
  }
  return groupNameCache.get(groupId);
}

const membersCache = new Map(); // groupId -> Promise<Map<userId, {name, mail, upn}>>
function groupMembers(groupId) {
  if (!groupId) return Promise.resolve(new Map());
  if (!membersCache.has(groupId)) {
    const p = getAll(`groups/${enc(groupId)}/members`, { $select: "id,displayName,mail,userPrincipalName", $top: 999 })
      .then((ms) => new Map(ms.map((m) => [m.id, { name: m.displayName, mail: m.mail, upn: m.userPrincipalName }])))
      .catch(() => new Map());
    membersCache.set(groupId, p);
    p.then(() => setTimeout(() => membersCache.delete(groupId), 5 * 60_000).unref());
  }
  return membersCache.get(groupId);
}

const userCache = new Map();
/**
 * Display name for a user id: group members first, then /users/{id}. Both need User.ReadBasic.All
 * (Group.Read.All alone returns members with null names); without it the id is returned.
 */
async function userName(id, members) {
  if (members?.get(id)?.name) return members.get(id).name;
  if (!userCache.has(id)) {
    userCache.set(
      id,
      request("GET", `users/${enc(id)}`, { params: { $select: "displayName" } })
        .then((u) => u.displayName)
        .catch(() => undefined),
    );
  }
  return (await userCache.get(id)) ?? id;
}

/** 'group' | 'roster' (Teams meeting / shared plans) | 'user' | ... from plan.container. */
const containerType = (plan) => {
  const c = plan.container;
  if (!c) return "group"; // old plans only have owner = group id
  if (c.type && c.type !== "unknownFutureValue") return c.type;
  return c.url?.match(/\/(roster|group|user|drive)s\//)?.[1] ?? c.type;
};
const groupIdOf = (plan) => (containerType(plan) === "group" ? (plan.container?.containerId ?? plan.owner) : undefined);

async function planContext(planId) {
  const id = enc(planId);
  const [plan, details, buckets] = await Promise.all([
    request("GET", `planner/plans/${id}`),
    request("GET", `planner/plans/${id}/details`),
    getAll(`planner/plans/${id}/buckets`),
  ]);
  const gid = groupIdOf(plan);
  const members = await groupMembers(gid);
  const labels = details.categoryDescriptions ?? {};
  return {
    plan,
    details,
    groupId: gid,
    members,
    buckets: boardOrder(buckets),
    bucketName: new Map(buckets.map((b) => [b.id, b.name])),
    labels, // { category1: 'Red label', ... }
  };
}

// The Planner board shows buckets left to right in DESCENDING ordinal orderHint (checked against the UI).
const boardOrder = (xs) => [...xs].sort((a, b) => (a.orderHint < b.orderHint ? 1 : a.orderHint > b.orderHint ? -1 : 0));

const categoryKeys = Array.from({ length: 25 }, (_, i) => `category${i + 1}`);
const categoriesOut = (applied, labels) =>
  Object.entries(applied ?? {})
    .filter(([, on]) => on)
    .map(([k]) => labels?.[k] || k);

/** Category label or 'categoryN' -> 'categoryN'. */
function categoryKey(name, labels) {
  if (/^category([1-9]|1\d|2[0-5])$/i.test(name)) return name.toLowerCase();
  const hit = categoryKeys.find((k) => (labels[k] ?? "").toLowerCase() === name.toLowerCase());
  if (hit) return hit;
  const known = categoryKeys.filter((k) => labels[k]).map((k) => `${labels[k]} (${k})`);
  throw new Error(`Unknown category "${name}". This plan's labels: ${known.join(", ") || "none named; use category1..category25"}.`);
}

/** User id, email/UPN, or 'me' -> user id. */
async function resolveUser(who, members) {
  if (isGuid(who)) return who;
  if (who.toLowerCase() === "me") return (await request("GET", "me", { params: { $select: "id" } })).id;
  const lower = who.toLowerCase();
  for (const [id, m] of members) {
    if (m.mail?.toLowerCase() === lower || m.upn?.toLowerCase() === lower) return id;
  }
  try {
    return (await request("GET", `users/${enc(who)}`, { params: { $select: "id" } })).id;
  } catch (e) {
    throw new Error(
      `Could not resolve "${who}" to a user: not a member of the plan's group, and the directory lookup failed (${e.message}). Pass the user's object id instead (email lookup needs User.ReadBasic.All: PLANNER_USER_NAMES in auth.js).`,
    );
  }
}

const assignment = () => ({ "@odata.type": "#microsoft.graph.plannerAssignment", orderHint: " !" });

// ---------- output shapes ----------

async function taskBrief(t, ctx) {
  const assignees = await Promise.all(Object.keys(t.assignments ?? {}).map((id) => userName(id, ctx?.members)));
  return {
    id: t.id,
    title: t.title,
    plan: ctx?.plan?.title,
    bucket: ctx?.bucketName?.get(t.bucketId) ?? t.bucketId,
    due: utcToDate(t.dueDateTime),
    start: utcToDate(t.startDateTime),
    percentComplete: t.percentComplete,
    status: statusOf(t.percentComplete),
    priority: priorityOut(t.priority),
    assignees,
    categories: categoriesOut(t.appliedCategories, ctx?.labels),
    checklist: t.checklistItemCount ? `${t.checklistItemCount - t.activeChecklistItemCount}/${t.checklistItemCount} done` : undefined,
    hasDescription: t.hasDescription || undefined,
  };
}

async function taskFull(t, d, ctx) {
  const assignees = await Promise.all(
    Object.keys(t.assignments ?? {}).map(async (id) => ({ id, name: await userName(id, ctx.members) })),
  );
  const checklist = Object.entries(d.checklist ?? {})
    .map(([id, c]) => ({ id, title: c.title, checked: c.isChecked, orderHint: c.orderHint }))
    .sort((a, b) => (a.orderHint < b.orderHint ? -1 : a.orderHint > b.orderHint ? 1 : 0))
    .map(({ orderHint, ...c }) => c);
  const references = Object.entries(d.references ?? {}).map(([url, r]) => ({
    url: decodeURIComponent(url),
    alias: r.alias || undefined,
    type: r.type || undefined,
  }));
  return {
    id: t.id,
    title: t.title,
    planId: t.planId,
    plan: ctx.plan.title,
    bucketId: t.bucketId,
    bucket: ctx.bucketName.get(t.bucketId),
    due: utcToDate(t.dueDateTime),
    start: utcToDate(t.startDateTime),
    dueUtc: t.dueDateTime ?? undefined,
    startUtc: t.startDateTime ?? undefined,
    dateTimeZone: PLANNER_TZ,
    percentComplete: t.percentComplete,
    status: statusOf(t.percentComplete),
    priority: priorityOut(t.priority),
    assignees,
    categories: categoriesOut(t.appliedCategories, ctx.labels),
    description: d.description || "",
    checklist,
    references,
    created: t.createdDateTime,
    createdBy: t.createdBy?.user?.id ? await userName(t.createdBy.user.id, ctx.members) : undefined,
    completed: t.completedDateTime ?? undefined,
    completedBy: t.completedBy?.user?.id ? await userName(t.completedBy.user.id, ctx.members) : undefined,
  };
}

async function readTask(taskId) {
  const id = enc(taskId);
  const [t, d] = await Promise.all([request("GET", `planner/tasks/${id}`), request("GET", `planner/tasks/${id}/details`)]);
  return taskFull(t, d, await planContext(t.planId));
}

const bucketIdFrom = (input, ctx) => {
  if (ctx.bucketName.has(input)) return input;
  const hit = ctx.buckets.find((b) => b.name.toLowerCase() === input.toLowerCase());
  if (hit) return hit.id;
  throw new Error(`No bucket "${input}" in plan "${ctx.plan.title}". Buckets: ${ctx.buckets.map((b) => b.name).join(", ")}.`);
};

// ---------- checklist helpers ----------

const ChecklistItem = z.union([z.string(), z.object({ title: z.string(), checked: z.boolean().default(false) })]);

const maxHint = (checklist) =>
  Object.values(checklist ?? {})
    .map((c) => c.orderHint)
    .sort()
    .at(-1);

/** New checklist entries (GUID keys) appended after the current last item. */
function newChecklistEntries(items, existing) {
  const last = maxHint(existing);
  const out = {};
  for (const it of items) {
    const { title, checked } = typeof it === "string" ? { title: it, checked: false } : it;
    out[randomUUID()] = {
      "@odata.type": "#microsoft.graph.plannerChecklistItem",
      title,
      isChecked: checked,
      orderHint: last ? `${last} !` : " !",
    };
  }
  return out;
}

const checklistKey = (ref, checklist) => {
  if (checklist?.[ref]) return ref;
  const hits = Object.entries(checklist ?? {}).filter(([, c]) => c.title.toLowerCase() === ref.toLowerCase());
  if (hits.length === 1) return hits[0][0];
  throw new Error(hits.length ? `Several checklist items are titled "${ref}"; use the item id.` : `No checklist item "${ref}".`);
};

// ---------- schemas ----------

const PlanParam = z.string().describe("Plan id, or a full Planner plan URL (the id is the segment after /plan/).");
const TaskParam = z.string().describe("Task id (or a Planner task URL).");
const DateParam = z
  .string()
  .regex(YMD, "YYYY-MM-DD")
  .describe("YYYY-MM-DD. Stored as 12:00 UTC that day, so it shows as that date in any time zone from UTC-11 to UTC+11.");
const UserList = z.array(z.string()).describe("User object ids, emails/UPNs, or 'me'.");

export function register(server) {
  server.tool(
    "planner_list_plans",
    "List Planner plans: yours, or one Microsoft 365 group's plans when group_id is given. Returns id, title, container type (group / roster) and owner group. Graph's /me/planner/plans omits most group plans, so 'yours' also includes the plans of your groups (when the directory permission allows listing them) and every plan that has a task assigned to you; a group plan with none of those can still be listed with group_id or opened by id/URL.",
    { group_id: z.string().optional().describe("Microsoft 365 group id. Omitted = your plans.") },
    async ({ group_id }) => {
      let plans;
      const notes = [];
      if (group_id) plans = await getAll(`groups/${enc(group_id)}/planner/plans`);
      else {
        const byId = new Map();
        const add = (ps) => ps.forEach((p) => p && byId.set(p.id, p));
        add(await getAll("me/planner/plans"));
        try {
          const groups = await getAll("me/memberOf/microsoft.graph.group", { $select: "id,groupTypes", $top: 999 });
          const unified = groups.filter((g) => g.groupTypes?.includes("Unified"));
          add((await Promise.all(unified.map((g) => getAll(`groups/${enc(g.id)}/planner/plans`).catch(() => [])))).flat());
        } catch (e) {
          notes.push(`Could not list your groups (${e.status ?? "error"}), so group plans without a task assigned to you may be missing; pass group_id to list a group's plans.`);
        }
        const myPlanIds = new Set((await getAll("me/planner/tasks")).map((t) => t.planId));
        const missing = [...myPlanIds].filter((id) => !byId.has(id));
        add(await Promise.all(missing.map((id) => request("GET", `planner/plans/${enc(id)}`).catch(() => undefined))));
        plans = [...byId.values()];
      }
      const out = await Promise.all(
        plans.map(async (p) => {
          const gid = groupIdOf(p);
          return {
            id: p.id,
            title: p.title,
            container: containerType(p),
            group: gid ? { id: gid, name: await groupName(gid) } : undefined,
            created: p.createdDateTime,
          };
        }),
      );
      out.sort((a, b) => (b.created ?? "").localeCompare(a.created ?? ""));
      return json(notes.length ? { plans: out, notes } : out);
    },
  );

  server.tool(
    "planner_get_plan",
    "Get a plan by id or Planner URL: title, owner group, category label names (category1..25) and buckets.",
    { plan_id: PlanParam },
    async ({ plan_id }) => {
      const ctx = await planContext(parsePlanId(plan_id));
      return json({
        id: ctx.plan.id,
        title: ctx.plan.title,
        container: containerType(ctx.plan),
        group: ctx.groupId ? { id: ctx.groupId, name: await groupName(ctx.groupId) } : undefined,
        created: ctx.plan.createdDateTime,
        categories: categoryKeys.filter((k) => ctx.labels[k]).map((k) => ({ key: k, label: ctx.labels[k] })),
        buckets: ctx.buckets.map((b) => ({ id: b.id, name: b.name })),
      });
    },
  );

  server.tool(
    "planner_list_buckets",
    "List the buckets in a plan, in board order (left to right).",
    { plan_id: PlanParam },
    async ({ plan_id }) => {
      const buckets = boardOrder(await getAll(`planner/plans/${enc(parsePlanId(plan_id))}/buckets`));
      return json(buckets.map((b) => ({ id: b.id, name: b.name, orderHint: b.orderHint })));
    },
  );

  server.tool(
    "planner_list_tasks",
    "List tasks in a plan with bucket names, due/start dates, progress, priority, assignee names and category labels. Open tasks only unless include_completed. Reads every page.",
    {
      plan_id: PlanParam,
      bucket_id: z.string().optional().describe("Only this bucket (id or name)."),
      include_completed: z.boolean().default(false),
    },
    async ({ plan_id, bucket_id, include_completed }) => {
      const pid = parsePlanId(plan_id);
      const [ctx, tasks] = await Promise.all([planContext(pid), getAll(`planner/plans/${enc(pid)}/tasks`)]);
      const bucket = bucket_id ? bucketIdFrom(bucket_id, ctx) : undefined;
      const picked = tasks.filter((t) => (include_completed || t.percentComplete < 100) && (!bucket || t.bucketId === bucket));
      const order = new Map(ctx.buckets.map((b, i) => [b.id, i]));
      picked.sort(
        (a, b) =>
          (order.get(a.bucketId) ?? 999) - (order.get(b.bucketId) ?? 999) ||
          (a.dueDateTime ?? "9999").localeCompare(b.dueDateTime ?? "9999"),
      );
      return json({
        plan: ctx.plan.title,
        dateTimeZone: PLANNER_TZ,
        count: picked.length,
        tasks: await Promise.all(picked.map((t) => taskBrief(t, ctx))),
      });
    },
  );

  server.tool(
    "planner_list_my_tasks",
    "List Planner tasks assigned to you across all plans, with plan and bucket names. Open tasks only unless include_completed.",
    { include_completed: z.boolean().default(false) },
    async ({ include_completed }) => {
      const tasks = (await getAll("me/planner/tasks")).filter((t) => include_completed || t.percentComplete < 100);
      const ctxs = new Map();
      for (const pid of new Set(tasks.map((t) => t.planId))) ctxs.set(pid, planContext(pid).catch(() => undefined));
      const out = await Promise.all(tasks.map(async (t) => taskBrief(t, await ctxs.get(t.planId))));
      out.sort((a, b) => (a.due ?? "9999").localeCompare(b.due ?? "9999"));
      return json({ dateTimeZone: PLANNER_TZ, count: out.length, tasks: out });
    },
  );

  server.tool(
    "planner_get_task",
    "Get one task with its details: description, checklist (with item ids), references/links, assignee names, category labels, dates and progress.",
    { task_id: TaskParam },
    async ({ task_id }) => json(await readTask(parseTaskId(task_id))),
  );

  server.tool(
    "planner_create_task",
    "Create a task in a plan. Description and checklist are written to the task's details right after it is created.",
    {
      plan_id: PlanParam,
      bucket_id: z.string().describe("Bucket id or bucket name."),
      title: z.string().min(1),
      due_date: DateParam.optional(),
      start_date: DateParam.optional(),
      percent_complete: PercentParam.optional(),
      priority: PriorityParam.optional(),
      assignees: UserList.optional(),
      categories: z.array(z.string()).optional().describe("Category label names from the plan (see planner_get_plan), or category1..category25."),
      description: z.string().optional().describe("Plain text."),
      checklist: z.array(ChecklistItem).max(20).optional().describe("Item titles, or { title, checked }. Planner allows 20 items."),
    },
    async (a) => {
      const ctx = await planContext(parsePlanId(a.plan_id));
      const body = { planId: ctx.plan.id, bucketId: bucketIdFrom(a.bucket_id, ctx), title: a.title };
      if (a.due_date) body.dueDateTime = dateToUtc(a.due_date);
      if (a.start_date) body.startDateTime = dateToUtc(a.start_date);
      if (a.percent_complete !== undefined) body.percentComplete = percentIn(a.percent_complete);
      if (a.priority !== undefined) body.priority = priorityIn(a.priority);
      if (a.assignees?.length) {
        body.assignments = {};
        for (const who of a.assignees) body.assignments[await resolveUser(who, ctx.members)] = assignment();
      }
      if (a.categories?.length) {
        body.appliedCategories = Object.fromEntries(a.categories.map((c) => [categoryKey(c, ctx.labels), true]));
      }

      const t = await request("POST", "planner/tasks", { json: body, headers: REPRESENTATION });

      let detailsError;
      if (a.description !== undefined || a.checklist?.length) {
        try {
          await waitForDetails(t.id);
          await patchWithEtag(
            `planner/tasks/${enc(t.id)}/details`,
            (d) => ({
              ...(a.description !== undefined && { description: a.description, previewType: "automatic" }),
              ...(a.checklist?.length && { checklist: newChecklistEntries(a.checklist, d.checklist) }),
            }),
            "task details",
          );
        } catch (e) {
          detailsError = `Task was created, but its description/checklist could not be saved: ${e.message}`;
        }
      }
      const full = await readTask(t.id);
      return json(detailsError ? { ...full, warning: detailsError } : full);
    },
  );

  server.tool(
    "planner_update_task",
    "Change a task. Only fields you pass change. Etags are handled for you (fetched before each PATCH; one retry on a 412/409 conflict). Clear a date with null. assignees/categories replace the whole set; the add_/remove_ variants change it incrementally.",
    {
      task_id: TaskParam,
      title: z.string().min(1).optional(),
      bucket_id: z.string().optional().describe("Move to this bucket (id or name, same plan)."),
      due_date: DateParam.nullable().optional(),
      start_date: DateParam.nullable().optional(),
      percent_complete: PercentParam.optional(),
      priority: PriorityParam.optional(),
      assignees: UserList.optional().describe("Replace assignees with exactly these (ids, emails or 'me'). [] unassigns everyone."),
      add_assignees: UserList.optional(),
      remove_assignees: UserList.optional(),
      categories: z.array(z.string()).optional().describe("Replace applied categories with exactly these (label names or categoryN)."),
      add_categories: z.array(z.string()).optional(),
      remove_categories: z.array(z.string()).optional(),
      description: z.string().optional().describe("Plain text; replaces the description. '' clears it."),
      checklist_add: z.array(ChecklistItem).optional().describe("Items to append (titles or { title, checked })."),
      checklist_update: z
        .array(z.object({ item: z.string().describe("Checklist item id or exact title."), title: z.string().optional(), checked: z.boolean().optional() }))
        .optional(),
      checklist_remove: z.array(z.string()).optional().describe("Checklist item ids or exact titles to remove."),
    },
    async (a) => {
      const taskId = parseTaskId(a.task_id);
      const id = enc(taskId);
      const head = await request("GET", `planner/tasks/${id}`);
      const ctx = await planContext(head.planId);

      // Resolve names/emails once, outside the retry loop.
      const ids = async (list) => (list ? Promise.all(list.map((w) => resolveUser(w, ctx.members))) : undefined);
      const keys = (list) => list?.map((c) => categoryKey(c, ctx.labels));
      const setAssignees = await ids(a.assignees);
      const addAssignees = await ids(a.add_assignees);
      const removeAssignees = await ids(a.remove_assignees);
      const setCats = keys(a.categories);
      const addCats = keys(a.add_categories);
      const removeCats = keys(a.remove_categories);
      const bucketId = a.bucket_id ? bucketIdFrom(a.bucket_id, ctx) : undefined;

      const taskFields = [
        "title", "bucket_id", "due_date", "start_date", "percent_complete", "priority",
        "assignees", "add_assignees", "remove_assignees", "categories", "add_categories", "remove_categories",
      ];
      const detailFields = ["description", "checklist_add", "checklist_update", "checklist_remove"];
      const touchesTask = taskFields.some((f) => a[f] !== undefined);
      const touchesDetails = detailFields.some((f) => a[f] !== undefined);
      if (!touchesTask && !touchesDetails) throw new Error("Nothing to update: pass at least one field.");

      if (touchesTask) {
        await patchWithEtag(
          `planner/tasks/${id}`,
          (t) => {
            const p = {};
            if (a.title !== undefined) p.title = a.title;
            if (bucketId) p.bucketId = bucketId;
            if (a.due_date !== undefined) p.dueDateTime = a.due_date === null ? null : dateToUtc(a.due_date);
            if (a.start_date !== undefined) p.startDateTime = a.start_date === null ? null : dateToUtc(a.start_date);
            if (a.percent_complete !== undefined) p.percentComplete = percentIn(a.percent_complete);
            if (a.priority !== undefined) p.priority = priorityIn(a.priority);

            const current = new Set(Object.keys(t.assignments ?? {}));
            const want = new Set(setAssignees ?? current);
            for (const u of addAssignees ?? []) want.add(u);
            for (const u of removeAssignees ?? []) want.delete(u);
            const asg = {};
            for (const u of want) if (!current.has(u)) asg[u] = assignment();
            for (const u of current) if (!want.has(u)) asg[u] = null; // null removes an assignee
            if (Object.keys(asg).length) p.assignments = asg;

            const applied = new Set(Object.entries(t.appliedCategories ?? {}).filter(([, v]) => v).map(([k]) => k));
            const cats = new Set(setCats ?? applied);
            for (const c of addCats ?? []) cats.add(c);
            for (const c of removeCats ?? []) cats.delete(c);
            const cat = {};
            for (const c of cats) if (!applied.has(c)) cat[c] = true;
            for (const c of applied) if (!cats.has(c)) cat[c] = false;
            if (Object.keys(cat).length) p.appliedCategories = cat;
            return p;
          },
          "task",
        );
      }

      if (touchesDetails) {
        try {
          await patchWithEtag(
            `planner/tasks/${id}/details`,
            (d) => {
              const p = {};
              if (a.description !== undefined) Object.assign(p, { description: a.description, previewType: "automatic" });
              const cl = {};
              for (const u of a.checklist_update ?? []) {
                const k = checklistKey(u.item, d.checklist);
                cl[k] = {
                  "@odata.type": "#microsoft.graph.plannerChecklistItem",
                  ...(u.title !== undefined && { title: u.title }),
                  ...(u.checked !== undefined && { isChecked: u.checked }),
                };
              }
              for (const r of a.checklist_remove ?? []) cl[checklistKey(r, d.checklist)] = null;
              if (a.checklist_add?.length) {
                const total = Object.keys(d.checklist ?? {}).length - (a.checklist_remove?.length ?? 0) + a.checklist_add.length;
                if (total > 20) throw new Error(`Planner allows 20 checklist items per task; this would make ${total}.`);
                Object.assign(cl, newChecklistEntries(a.checklist_add, d.checklist));
              }
              if (Object.keys(cl).length) p.checklist = cl;
              return p;
            },
            "task details",
          );
        } catch (e) {
          if (touchesTask) e.message = `Task fields were saved, but the details (description/checklist) were not: ${e.message}`;
          throw e;
        }
      }

      return json(await readTask(taskId));
    },
  );

  server.tool(
    "planner_complete_task",
    "Mark a task completed (percentComplete 100), or reopen it with completed=false (back to 0, not started).",
    { task_id: TaskParam, completed: z.boolean().default(true) },
    async ({ task_id, completed }) => {
      const taskId = parseTaskId(task_id);
      await patchWithEtag(
        `planner/tasks/${enc(taskId)}`,
        (t) => {
          const target = completed ? 100 : 0;
          return t.percentComplete === target ? null : { percentComplete: target };
        },
        "task",
      );
      return json(await readTask(taskId));
    },
  );
}
