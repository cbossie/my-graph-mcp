// Calendar tools.
//
// Deliberately NOT here: sending invites and responding to invites. Microsoft Graph has no draft
// event, so creating or changing an event that has attendees emails them immediately. To keep
// that off the table the server enforces these rules (Calendars.ReadWrite itself cannot):
//   - no tool accepts attendees, and there are no accept/decline/tentative tools;
//   - updating the substance of an event, or deleting it, is refused when anyone other than you
//     is on it (the change or cancellation would be mailed out).
// To invite people, create the event here and add attendees in Outlook, where you press Send.
import { z } from "zod";
import { request } from "../graph.js";

const json = (value) => ({ content: [{ type: "text", text: JSON.stringify(value, null, 2) }] });

const DEFAULT_TZ = process.env.CALENDAR_TIME_ZONE || Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";

const TimeZone = z
  .string()
  .optional()
  .describe(`Time zone for start/end and for results (IANA like 'America/New_York' or Windows name). Default: ${DEFAULT_TZ}.`);
const CalendarId = z.string().optional().describe("Calendar id (from calendar_list_calendars). Omitted = your default calendar.");
const SHOW_AS = ["free", "tentative", "busy", "oof", "workingElsewhere"];
const DAYS = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];

const Recurrence = z
  .object({
    type: z.enum(["daily", "weekly", "absoluteMonthly", "relativeMonthly", "absoluteYearly", "relativeYearly"]),
    interval: z.number().int().min(1).default(1).describe("Every N days/weeks/months/years."),
    days_of_week: z.array(z.enum(DAYS)).optional().describe("weekly, relativeMonthly, relativeYearly."),
    day_of_month: z.number().int().min(1).max(31).optional().describe("absoluteMonthly, absoluteYearly."),
    month: z.number().int().min(1).max(12).optional().describe("absoluteYearly, relativeYearly."),
    index: z.enum(["first", "second", "third", "fourth", "last"]).optional().describe("relativeMonthly, relativeYearly."),
    first_day_of_week: z.enum(DAYS).optional(),
    ends: z
      .object({
        type: z.enum(["noEnd", "endDate", "numbered"]),
        end_date: z.string().optional().describe("YYYY-MM-DD, when type is endDate."),
        occurrences: z.number().int().min(1).optional().describe("When type is numbered."),
      })
      .default({ type: "noEnd" }),
  })
  .describe("Repeat rule; the series starts on the event's start date.");

const dateOnly = (s) => s.slice(0, 10);
const addDays = (ymd, n) => {
  const d = new Date(`${ymd}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
};

/** Build Graph start/end. All-day events take dates; `end` is the LAST day (inclusive). */
function timing({ start, end, all_day, time_zone }) {
  const timeZone = time_zone ?? DEFAULT_TZ;
  if (all_day) {
    const s = dateOnly(start);
    const e = dateOnly(end ?? start);
    return {
      isAllDay: true,
      start: { dateTime: `${s}T00:00:00`, timeZone },
      end: { dateTime: `${addDays(e, 1)}T00:00:00`, timeZone },
    };
  }
  return { start: { dateTime: start, timeZone }, end: { dateTime: end, timeZone } };
}

function recurrenceBody(r, startDate, time_zone) {
  const pattern = { type: r.type, interval: r.interval };
  if (r.days_of_week) pattern.daysOfWeek = r.days_of_week;
  if (r.day_of_month) pattern.dayOfMonth = r.day_of_month;
  if (r.month) pattern.month = r.month;
  if (r.index) pattern.index = r.index;
  if (r.first_day_of_week) pattern.firstDayOfWeek = r.first_day_of_week;
  const range = { type: r.ends.type, startDate, recurrenceTimeZone: time_zone ?? DEFAULT_TZ };
  if (r.ends.type === "endDate") range.endDate = r.ends.end_date;
  if (r.ends.type === "numbered") range.numberOfOccurrences = r.ends.occurrences;
  return { pattern, range };
}

const eventPath = (calendarId) => (calendarId ? `me/calendars/${encodeURIComponent(calendarId)}/events` : "me/events");
const tzHeader = (tz) => ({ Prefer: `outlook.timezone="${tz ?? DEFAULT_TZ}"` });

const SELECT =
  "id,subject,start,end,isAllDay,location,organizer,attendees,isOnlineMeeting,onlineMeeting,showAs,importance,sensitivity,categories,type,seriesMasterId,recurrence,isCancelled,isOrganizer,responseStatus,reminderMinutesBeforeStart,webLink";

const summary = (e, tz) => ({
  id: e.id,
  subject: e.subject,
  start: e.start?.dateTime,
  end: e.end?.dateTime,
  timeZone: e.start?.timeZone ?? tz,
  allDay: e.isAllDay,
  location: e.location?.displayName || undefined,
  organizer: e.organizer?.emailAddress?.name,
  isOrganizer: e.isOrganizer,
  attendees: (e.attendees ?? []).map((a) => ({
    name: a.emailAddress?.name,
    email: a.emailAddress?.address,
    type: a.type,
    response: a.status?.response,
  })),
  myResponse: e.responseStatus?.response,
  showAs: e.showAs,
  categories: e.categories?.length ? e.categories : undefined,
  type: e.type, // singleInstance | occurrence | exception | seriesMaster
  seriesMasterId: e.seriesMasterId || undefined,
  recurrence: e.recurrence || undefined,
  onlineMeetingUrl: e.onlineMeeting?.joinUrl || undefined,
  reminderMinutes: e.reminderMinutesBeforeStart,
  cancelled: e.isCancelled || undefined,
  webLink: e.webLink,
});

let mine;
async function myAddresses() {
  if (!mine) {
    try {
      const me = await request("GET", "me", { params: { $select: "mail,userPrincipalName" } });
      mine = new Set([me.mail, me.userPrincipalName].filter(Boolean).map((a) => a.toLowerCase()));
    } catch {
      mine = new Set(); // can't tell who "me" is: treat every attendee as someone else (safer)
    }
  }
  return mine;
}

/** Attendees other than you. Changing or deleting such an event would email them. */
async function otherAttendees(event) {
  const me = await myAddresses();
  return (event.attendees ?? []).filter((a) => !me.has(a.emailAddress?.address?.toLowerCase()));
}

const refuse = (others, action) =>
  new Error(
    `Refusing to ${action}: this event has other attendees (${others
      .slice(0, 5)
      .map((a) => a.emailAddress?.name || a.emailAddress?.address)
      .join(", ")}${others.length > 5 ? ", ..." : ""}) and Outlook would email them. This server never sends invites, updates or cancellations. Do it in Outlook instead.`,
  );

/** Compact form for lists: no emails/URLs/links; use calendar_get_event for the full detail. */
const brief = (e, tz) => {
  const s = summary(e, tz);
  const people = s.attendees.filter((a) => a.email);
  return {
    id: s.id,
    subject: s.subject,
    start: s.start,
    end: s.end,
    allDay: s.allDay || undefined,
    location: s.location,
    organizer: s.isOrganizer ? "me" : s.organizer,
    attendees: people.length > 1 ? `${people.length} people` : undefined,
    myResponse: s.myResponse,
    showAs: s.showAs,
    categories: s.categories,
    recurring: s.type === "singleInstance" ? undefined : s.type,
    online: s.onlineMeetingUrl ? true : undefined,
  };
};

export function register(server) {
  server.tool("calendar_list_calendars", "List your calendars (id, name, owner, whether you can edit).", {}, async () => {
    const r = await request("GET", "me/calendars", { params: { $select: "id,name,owner,canEdit,isDefaultCalendar" } });
    return json(
      r.value.map((c) => ({
        id: c.id,
        name: c.name,
        owner: c.owner?.address,
        canEdit: c.canEdit,
        isDefault: c.isDefaultCalendar || undefined,
      })),
    );
  });

  server.tool(
    "calendar_list_events",
    "List events between two times, with recurring series expanded into individual occurrences, sorted by start. Compact: attendee count and your response only; use calendar_get_event for attendees, links and the body.",
    {
      start: z.string().describe("Range start, local time 'YYYY-MM-DDTHH:mm:ss' (or just a date)."),
      end: z.string().describe("Range end, same format."),
      time_zone: TimeZone,
      calendar_id: CalendarId,
      top: z.number().int().min(1).max(100).default(50),
    },
    async ({ start, end, time_zone, calendar_id, top }) => {
      const base = calendar_id ? `me/calendars/${encodeURIComponent(calendar_id)}/calendarView` : "me/calendarView";
      const toLocal = (s) => (s.length === 10 ? `${s}T00:00:00` : s);
      const r = await request("GET", base, {
        params: {
          startDateTime: toLocal(start),
          endDateTime: toLocal(end),
          $select: SELECT,
          $orderby: "start/dateTime",
          $top: top,
        },
        headers: tzHeader(time_zone),
      });
      return json({ events: r.value.map((e) => brief(e, time_zone)), more: Boolean(r["@odata.nextLink"]) });
    },
  );

  server.tool(
    "calendar_get_event",
    "Get one event in full, including its body as text.",
    { event_id: z.string(), time_zone: TimeZone },
    async ({ event_id, time_zone }) => {
      const e = await request("GET", `me/events/${encodeURIComponent(event_id)}`, {
        params: { $select: `${SELECT},body` },
        headers: { Prefer: `outlook.timezone="${time_zone ?? DEFAULT_TZ}", outlook.body-content-type="text"` },
      });
      return json({ ...summary(e, time_zone), body: e.body?.content });
    },
  );

  server.tool(
    "calendar_create_event",
    "Create an event on your own calendar. There is NO attendees option by design: nobody is invited or emailed. To invite people, add them in Outlook and press Send there. all_day events take dates; `end` is the last day (inclusive).",
    {
      subject: z.string(),
      start: z.string().describe("Local time 'YYYY-MM-DDTHH:mm:ss' (or a date for all_day)."),
      end: z.string().optional().describe("Local time; for all_day the last day (default: same day). Required otherwise."),
      all_day: z.boolean().default(false),
      time_zone: TimeZone,
      location: z.string().optional(),
      body_html: z.string().optional().describe("HTML description."),
      show_as: z.enum(SHOW_AS).optional(),
      importance: z.enum(["low", "normal", "high"]).optional(),
      sensitivity: z.enum(["normal", "personal", "private", "confidential"]).optional(),
      categories: z.array(z.string()).optional(),
      reminder_minutes: z.number().int().min(0).optional().describe("Reminder this many minutes before start."),
      online_meeting: z.boolean().default(false).describe("Add a Teams meeting link."),
      recurrence: Recurrence.optional(),
      calendar_id: CalendarId,
    },
    async (a) => {
      if (!a.all_day && !a.end) throw new Error("`end` is required unless all_day is true.");
      const body = {
        subject: a.subject,
        ...timing(a),
        ...(a.location && { location: { displayName: a.location } }),
        ...(a.body_html && { body: { contentType: "HTML", content: a.body_html } }),
        ...(a.show_as && { showAs: a.show_as }),
        ...(a.importance && { importance: a.importance }),
        ...(a.sensitivity && { sensitivity: a.sensitivity }),
        ...(a.categories && { categories: a.categories }),
        ...(a.reminder_minutes !== undefined && { isReminderOn: true, reminderMinutesBeforeStart: a.reminder_minutes }),
        ...(a.online_meeting && { isOnlineMeeting: true, onlineMeetingProvider: "teamsForBusiness" }),
        ...(a.recurrence && { recurrence: recurrenceBody(a.recurrence, dateOnly(a.start), a.time_zone) }),
      };
      const e = await request("POST", eventPath(a.calendar_id), { json: body, headers: tzHeader(a.time_zone) });
      return json(summary(e, a.time_zone));
    },
  );

  server.tool(
    "calendar_update_event",
    "Change an event. Only the fields you pass change. If the event has other attendees, only categories, reminder_minutes and show_as can be changed (anything else would email them, so it is refused). For a recurring series pass the series master id; pass an occurrence id to change one occurrence.",
    {
      event_id: z.string(),
      subject: z.string().optional(),
      start: z.string().optional(),
      end: z.string().optional(),
      all_day: z.boolean().optional().describe("Set together with start (and end) to switch an event to or from all-day."),
      time_zone: TimeZone,
      location: z.string().optional(),
      body_html: z.string().optional(),
      show_as: z.enum(SHOW_AS).optional(),
      importance: z.enum(["low", "normal", "high"]).optional(),
      sensitivity: z.enum(["normal", "personal", "private", "confidential"]).optional(),
      categories: z.array(z.string()).optional(),
      reminder_minutes: z.number().int().min(0).optional(),
      online_meeting: z.boolean().optional().describe("true adds a Teams link."),
      recurrence: Recurrence.optional().describe("Needs `start` to anchor the series start date."),
    },
    async (a) => {
      const id = encodeURIComponent(a.event_id);
      const current = await request("GET", `me/events/${id}`, { params: { $select: "id,attendees,isOrganizer,start" } });

      const patch = {};
      const quiet = {}; // fields that do not notify attendees
      if (a.categories) quiet.categories = a.categories;
      if (a.show_as) quiet.showAs = a.show_as;
      if (a.reminder_minutes !== undefined) Object.assign(quiet, { isReminderOn: true, reminderMinutesBeforeStart: a.reminder_minutes });

      if (a.subject !== undefined) patch.subject = a.subject;
      if (a.start || a.end) {
        if (a.all_day) Object.assign(patch, timing({ start: a.start, end: a.end, all_day: true, time_zone: a.time_zone }));
        else {
          const tz = a.time_zone ?? DEFAULT_TZ;
          if (a.start) patch.start = { dateTime: a.start, timeZone: tz };
          if (a.end) patch.end = { dateTime: a.end, timeZone: tz };
          if (a.all_day === false) patch.isAllDay = false;
        }
      }
      if (a.location !== undefined) patch.location = { displayName: a.location };
      if (a.body_html !== undefined) patch.body = { contentType: "HTML", content: a.body_html };
      if (a.importance) patch.importance = a.importance;
      if (a.sensitivity) patch.sensitivity = a.sensitivity;
      if (a.online_meeting) Object.assign(patch, { isOnlineMeeting: true, onlineMeetingProvider: "teamsForBusiness" });
      if (a.recurrence) {
        if (!a.start) throw new Error("recurrence needs `start` to set the series start date.");
        patch.recurrence = recurrenceBody(a.recurrence, dateOnly(a.start), a.time_zone);
      }

      if (Object.keys(patch).length) {
        const others = await otherAttendees(current);
        if (others.length) throw refuse(others, "change this event");
      }
      const body = { ...patch, ...quiet };
      if (!Object.keys(body).length) throw new Error("Nothing to update: pass at least one field.");

      const e = await request("PATCH", `me/events/${id}`, { json: body, headers: tzHeader(a.time_zone) });
      return json(summary(e, a.time_zone));
    },
  );

  server.tool(
    "calendar_delete_event",
    "Delete an event from your calendar (it goes to Deleted Items). Refused if the event has other attendees, since Outlook would send them a cancellation. Show the user which event will be deleted and get their OK before calling with confirm=true.",
    { event_id: z.string(), confirm: z.literal(true) },
    async ({ event_id }) => {
      const id = encodeURIComponent(event_id);
      const current = await request("GET", `me/events/${id}`, { params: { $select: "id,subject,attendees" } });
      const others = await otherAttendees(current);
      if (others.length) throw refuse(others, "delete this event");
      await request("DELETE", `me/events/${id}`);
      return json({ deleted: event_id, subject: current.subject });
    },
  );
}
