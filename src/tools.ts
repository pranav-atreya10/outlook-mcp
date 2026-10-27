// The tools Claude can use: Outlook email + Outlook calendar.
import { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { graph } from "./microsoft";

export interface Props extends Record<string, unknown> {
  email: string;
  name: string;
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
}

type Result = { content: { type: "text"; text: string }[]; isError?: boolean };

const ok = (data: unknown): Result => ({
  content: [{ type: "text", text: typeof data === "string" ? data : JSON.stringify(data, null, 2) }],
});

// Wrap every tool so a Microsoft error becomes a readable message instead of a crash.
function safe<A>(fn: (args: A) => Promise<Result>) {
  return async (args: A): Promise<Result> => {
    try {
      return await fn(args);
    } catch (e) {
      return { content: [{ type: "text", text: e instanceof Error ? e.message : String(e) }], isError: true };
    }
  };
}

const recipients = (list: string[] = []) => list.map((address) => ({ emailAddress: { address } }));

// Keep email results small so Claude isn't flooded.
function summarizeMessage(m: any) {
  return {
    id: m.id,
    subject: m.subject,
    from: m.from?.emailAddress ? `${m.from.emailAddress.name} <${m.from.emailAddress.address}>` : undefined,
    received: m.receivedDateTime,
    isRead: m.isRead,
    hasAttachments: m.hasAttachments,
    preview: m.bodyPreview,
  };
}

function summarizeEvent(e: any) {
  return {
    id: e.id,
    subject: e.subject,
    start: e.start,
    end: e.end,
    isAllDay: e.isAllDay,
    location: e.location?.displayName || undefined,
    organizer: e.organizer?.emailAddress?.address,
    attendees: (e.attendees ?? []).map((a: any) => `${a.emailAddress?.address} (${a.status?.response ?? "none"})`),
    onlineMeetingUrl: e.onlineMeeting?.joinUrl ?? undefined,
    preview: e.bodyPreview || undefined,
    // "occurrence" = one date of a repeating event; seriesId points at the whole series
    repeats: e.type && e.type !== "singleInstance" ? e.type : undefined,
    seriesId: e.seriesMasterId ?? (e.type === "seriesMaster" ? e.id : undefined),
  };
}

const DAYS = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"] as const;

const recurrenceSchema = z
  .object({
    frequency: z.enum(["daily", "weekly", "monthly"]),
    interval: z.number().int().min(1).default(1).describe("Every N days/weeks/months"),
    daysOfWeek: z.array(z.enum(DAYS)).optional().describe("For weekly: which days. Defaults to the start date's day."),
    endDate: z.string().optional().describe("Last date it can happen, like 2027-03-24. Leave out to repeat forever."),
    occurrences: z.number().int().min(1).optional().describe("Or: stop after this many times"),
  })
  .describe("Make the event repeat");

function toGraphRecurrence(r: z.infer<typeof recurrenceSchema>, startDateTime: string) {
  const startDate = startDateTime.slice(0, 10);
  const dayOfStart = DAYS[new Date(`${startDate}T12:00:00Z`).getUTCDay()];
  const pattern: Record<string, unknown> = { type: r.frequency === "monthly" ? "absoluteMonthly" : r.frequency, interval: r.interval };
  if (r.frequency === "weekly") pattern.daysOfWeek = r.daysOfWeek?.length ? r.daysOfWeek : [dayOfStart];
  if (r.frequency === "monthly") pattern.dayOfMonth = Number(startDate.slice(8, 10));
  const range = r.endDate
    ? { type: "endDate", startDate, endDate: r.endDate }
    : r.occurrences
      ? { type: "numbered", startDate, numberOfOccurrences: r.occurrences }
      : { type: "noEnd", startDate };
  return { pattern, range };
}

const folderEnum = z
  .string()
  .describe("Folder: inbox, sentitems, drafts, deleteditems, junkemail, archive, or a folder id")
  .default("inbox");

const dateTime = z
  .string()
  .describe("Local date and time, like 2026-10-30T15:00:00 (no Z). For all-day events use midnight.");

export function buildServer(props: Props, defaultTimeZone: string): McpServer {
  const server = new McpServer({ name: "outlook-personal", version: "1.0.0" });
  const token = props.accessToken;
  const tzHeader = (tz?: string) => ({ Prefer: `outlook.timezone="${tz ?? defaultTimeZone}"` });

  // ---------------- EMAIL ----------------

  server.registerTool(
    "whoami",
    {
      description: "Shows which Outlook account this connector is signed in to.",
      inputSchema: z.object({}),
      annotations: { readOnlyHint: true },
    },
    safe(async () => ok({ name: props.name, email: props.email, defaultTimeZone })),
  );

  server.registerTool(
    "search_emails",
    {
      description:
        "List or search Outlook emails. With no query, returns the newest emails in the folder. With a query, searches subject, body and sender (e.g. 'from:alex@x.com', 'subject:invoice', or plain words).",
      inputSchema: z.object({
        query: z.string().optional().describe("Search words. Leave empty to just list the newest emails."),
        folder: folderEnum,
        unreadOnly: z.boolean().default(false).describe("Only unread emails (ignored when a query is given)"),
        limit: z.number().int().min(1).max(50).default(15),
      }),
      annotations: { readOnlyHint: true },
    },
    safe(async ({ query, folder, unreadOnly, limit }) => {
      const select = "id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview";
      const q: Record<string, string | undefined> = { $top: String(limit), $select: select };
      if (query) {
        // Graph can't sort or filter together with $search.
        q.$search = `"${query.replace(/"/g, "")}"`;
      } else {
        q.$orderby = "receivedDateTime desc";
        if (unreadOnly) q.$filter = "isRead eq false";
      }
      const data = await graph(token, `/me/mailFolders/${encodeURIComponent(folder)}/messages`, { query: q });
      return ok((data.value ?? []).map(summarizeMessage));
    }),
  );

  server.registerTool(
    "read_email",
    {
      description: "Open one email and read its full text. Use an id from search_emails.",
      inputSchema: z.object({ id: z.string() }),
      annotations: { readOnlyHint: true },
    },
    safe(async ({ id }) => {
      const m = await graph(token, `/me/messages/${encodeURIComponent(id)}`, {
        query: {
          $select:
            "id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,hasAttachments,body,webLink",
        },
        headers: { Prefer: 'outlook.body-content-type="text"' },
      });
      const addr = (r: any[] = []) => r.map((x) => x.emailAddress?.address).join(", ");
      return ok({
        id: m.id,
        subject: m.subject,
        from: m.from?.emailAddress?.address,
        to: addr(m.toRecipients),
        cc: addr(m.ccRecipients),
        received: m.receivedDateTime,
        isRead: m.isRead,
        hasAttachments: m.hasAttachments,
        link: m.webLink,
        body: String(m.body?.content ?? "").slice(0, 20000),
      });
    }),
  );

  server.registerTool(
    "send_email",
    {
      description: "Send a new email from the signed-in Outlook account. Always confirm the recipients and text with the user first.",
      inputSchema: z.object({
        to: z.array(z.string().email()).min(1),
        cc: z.array(z.string().email()).optional(),
        bcc: z.array(z.string().email()).optional(),
        subject: z.string(),
        body: z.string(),
        isHtml: z.boolean().default(false),
      }),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    },
    safe(async ({ to, cc, bcc, subject, body, isHtml }) => {
      await graph(token, "/me/sendMail", {
        method: "POST",
        body: {
          message: {
            subject,
            body: { contentType: isHtml ? "HTML" : "Text", content: body },
            toRecipients: recipients(to),
            ccRecipients: recipients(cc),
            bccRecipients: recipients(bcc),
          },
          saveToSentItems: true,
        },
      });
      return ok(`Sent "${subject}" to ${to.join(", ")}.`);
    }),
  );

  server.registerTool(
    "draft_email",
    {
      description: "Save an email as a draft in Outlook without sending it.",
      inputSchema: z.object({
        to: z.array(z.string().email()).default([]),
        cc: z.array(z.string().email()).optional(),
        subject: z.string(),
        body: z.string(),
        isHtml: z.boolean().default(false),
      }),
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    safe(async ({ to, cc, subject, body, isHtml }) => {
      const m = await graph(token, "/me/messages", {
        method: "POST",
        body: {
          subject,
          body: { contentType: isHtml ? "HTML" : "Text", content: body },
          toRecipients: recipients(to),
          ccRecipients: recipients(cc),
        },
      });
      return ok({ saved: true, id: m.id, link: m.webLink });
    }),
  );

  server.registerTool(
    "reply_email",
    {
      description: "Reply to an email (or reply-all). Always confirm the reply text with the user first.",
      inputSchema: z.object({
        id: z.string().describe("Email id from search_emails"),
        body: z.string().describe("Your reply text"),
        replyAll: z.boolean().default(false),
      }),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    },
    safe(async ({ id, body, replyAll }) => {
      await graph(token, `/me/messages/${encodeURIComponent(id)}/${replyAll ? "replyAll" : "reply"}`, {
        method: "POST",
        body: { comment: body },
      });
      return ok(replyAll ? "Reply-all sent." : "Reply sent.");
    }),
  );

  server.registerTool(
    "forward_email",
    {
      description: "Forward an email to other people, with an optional note.",
      inputSchema: z.object({
        id: z.string(),
        to: z.array(z.string().email()).min(1),
        note: z.string().default(""),
      }),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    },
    safe(async ({ id, to, note }) => {
      await graph(token, `/me/messages/${encodeURIComponent(id)}/forward`, {
        method: "POST",
        body: { comment: note, toRecipients: recipients(to) },
      });
      return ok(`Forwarded to ${to.join(", ")}.`);
    }),
  );

  server.registerTool(
    "update_email",
    {
      description: "Mark an email read/unread, flag it, or move it to another folder (e.g. archive, deleteditems to delete it).",
      inputSchema: z.object({
        id: z.string(),
        isRead: z.boolean().optional(),
        flagged: z.boolean().optional(),
        moveTo: z.string().optional().describe("Folder to move it to: archive, deleteditems, junkemail, inbox, or a folder id"),
      }),
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    safe(async ({ id, isRead, flagged, moveTo }) => {
      const path = `/me/messages/${encodeURIComponent(id)}`;
      const patch: Record<string, unknown> = {};
      if (isRead !== undefined) patch.isRead = isRead;
      if (flagged !== undefined) patch.flag = { flagStatus: flagged ? "flagged" : "notFlagged" };
      if (Object.keys(patch).length) await graph(token, path, { method: "PATCH", body: patch });
      let newId = id;
      if (moveTo) newId = (await graph(token, `${path}/move`, { method: "POST", body: { destinationId: moveTo } })).id;
      return ok({ done: true, id: newId });
    }),
  );

  server.registerTool(
    "list_mail_folders",
    {
      description: "List the Outlook mail folders with unread counts.",
      inputSchema: z.object({}),
      annotations: { readOnlyHint: true },
    },
    safe(async () => {
      const data = await graph(token, "/me/mailFolders", {
        query: { $top: "100", $select: "id,displayName,unreadItemCount,totalItemCount" },
      });
      return ok(data.value);
    }),
  );

  // ---------------- CALENDAR ----------------

  server.registerTool(
    "list_events",
    {
      description: `List Outlook calendar events between two dates (recurring events are expanded). Times are in ${defaultTimeZone} unless timeZone is given.`,
      inputSchema: z.object({
        start: dateTime.describe("Start of the range, e.g. 2026-10-27T00:00:00"),
        end: dateTime.describe("End of the range, e.g. 2026-11-03T00:00:00"),
        timeZone: z.string().optional().describe("IANA time zone, e.g. America/New_York"),
        limit: z.number().int().min(1).max(100).default(50),
      }),
      annotations: { readOnlyHint: true },
    },
    safe(async ({ start, end, timeZone, limit }) => {
      const data = await graph(token, "/me/calendarView", {
        query: {
          startDateTime: start,
          endDateTime: end,
          $top: String(limit),
          $orderby: "start/dateTime",
          $select: "id,subject,start,end,isAllDay,location,organizer,attendees,onlineMeeting,bodyPreview,type,seriesMasterId",
        },
        headers: tzHeader(timeZone),
      });
      return ok((data.value ?? []).map(summarizeEvent));
    }),
  );

  server.registerTool(
    "get_event",
    {
      description: "Get full details for one calendar event.",
      inputSchema: z.object({ id: z.string(), timeZone: z.string().optional() }),
      annotations: { readOnlyHint: true },
    },
    safe(async ({ id, timeZone }) => {
      const e = await graph(token, `/me/events/${encodeURIComponent(id)}`, {
        headers: { ...tzHeader(timeZone), Prefer: `${tzHeader(timeZone).Prefer}, outlook.body-content-type="text"` },
      });
      return ok({ ...summarizeEvent(e), body: e.body?.content, recurrence: e.recurrence, link: e.webLink });
    }),
  );

  const eventFields = {
    subject: z.string(),
    start: dateTime,
    end: dateTime,
    timeZone: z.string().optional().describe(`IANA time zone. Default: ${defaultTimeZone}`),
    isAllDay: z.boolean().optional(),
    location: z.string().optional(),
    description: z.string().optional(),
    attendees: z.array(z.string().email()).optional().describe("Invitees. Outlook emails them an invite."),
    reminderMinutes: z.number().int().min(0).optional(),
  };

  function eventBody(a: Partial<z.infer<z.ZodObject<typeof eventFields>>>) {
    const tz = a.timeZone ?? defaultTimeZone;
    const b: Record<string, unknown> = {};
    if (a.subject !== undefined) b.subject = a.subject;
    if (a.start !== undefined) b.start = { dateTime: a.start, timeZone: tz };
    if (a.end !== undefined) b.end = { dateTime: a.end, timeZone: tz };
    if (a.isAllDay !== undefined) b.isAllDay = a.isAllDay;
    if (a.location !== undefined) b.location = { displayName: a.location };
    if (a.description !== undefined) b.body = { contentType: "Text", content: a.description };
    if (a.attendees !== undefined)
      b.attendees = a.attendees.map((address) => ({ emailAddress: { address }, type: "required" }));
    if (a.reminderMinutes !== undefined) {
      b.isReminderOn = true;
      b.reminderMinutesBeforeStart = a.reminderMinutes;
    }
    return b;
  }

  server.registerTool(
    "create_event",
    {
      description:
        "Add an event to the Outlook calendar. Use `repeat` for a repeating event (one series, one invite). If attendees are given, they get an invite email.",
      inputSchema: z.object({ ...eventFields, repeat: recurrenceSchema.optional() }),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    },
    safe(async ({ repeat, ...args }) => {
      const body = eventBody(args);
      if (repeat) body.recurrence = toGraphRecurrence(repeat, args.start);
      const e = await graph(token, "/me/events", { method: "POST", body, headers: tzHeader(args.timeZone) });
      return ok({ created: true, ...summarizeEvent(e), link: e.webLink });
    }),
  );

  server.registerTool(
    "update_event",
    {
      description: "Change an existing Outlook calendar event. Only send the fields you want to change.",
      inputSchema: z.object({
        id: z.string(),
        subject: z.string().optional(),
        start: dateTime.optional(),
        end: dateTime.optional(),
        timeZone: z.string().optional(),
        isAllDay: z.boolean().optional(),
        location: z.string().optional(),
        description: z.string().optional(),
        attendees: z.array(z.string().email()).optional().describe("Replaces the whole attendee list"),
        reminderMinutes: z.number().int().min(0).optional(),
      }),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
    },
    safe(async ({ id, ...rest }) => {
      const e = await graph(token, `/me/events/${encodeURIComponent(id)}`, {
        method: "PATCH",
        body: eventBody(rest),
        headers: tzHeader(rest.timeZone),
      });
      return ok({ updated: true, ...summarizeEvent(e) });
    }),
  );

  server.registerTool(
    "delete_event",
    {
      description:
        "Delete an event from the Outlook calendar. For a repeating event, deletes just that one date unless entireSeries is true. If you organized it with attendees, they get a cancellation. Confirm with the user first.",
      inputSchema: z.object({
        id: z.string(),
        entireSeries: z.boolean().default(false).describe("For repeating events: delete every date, past and future"),
        cancelMessage: z.string().optional().describe("Optional note sent to attendees"),
      }),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
    },
    safe(async ({ id, entireSeries, cancelMessage }) => {
      let target = id;
      if (entireSeries) {
        const e = await graph(token, `/me/events/${encodeURIComponent(id)}`, { query: { $select: "seriesMasterId" } });
        target = e.seriesMasterId ?? id;
      }
      const path = `/me/events/${encodeURIComponent(target)}`;
      if (cancelMessage) {
        // cancel only works for meetings you organized; fall back to delete otherwise
        try {
          await graph(token, `${path}/cancel`, { method: "POST", body: { comment: cancelMessage } });
          return ok("Event cancelled and attendees notified.");
        } catch {
          /* not the organizer: just delete */
        }
      }
      await graph(token, path, { method: "DELETE" });
      return ok(entireSeries ? "Whole series deleted." : "Event deleted.");
    }),
  );

  server.registerTool(
    "stop_repeating_event",
    {
      description:
        "End a repeating event after a given date: keeps past dates, removes all later ones. Pass any date's id from the series. Confirm with the user first.",
      inputSchema: z.object({
        id: z.string().describe("Id of any date in the series (from list_events)"),
        lastDate: z.string().describe("Last date to keep, like 2026-10-24. Dates after this are removed."),
      }),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
    },
    safe(async ({ id, lastDate }) => {
      const e = await graph(token, `/me/events/${encodeURIComponent(id)}`, { query: { $select: "seriesMasterId,type" } });
      const masterId = e.seriesMasterId ?? (e.type === "seriesMaster" ? id : undefined);
      if (!masterId) return { content: [{ type: "text", text: "That event doesn't repeat." }], isError: true };
      const master = await graph(token, `/me/events/${encodeURIComponent(masterId)}`, { query: { $select: "subject,recurrence" } });
      const { pattern, range } = master.recurrence;
      if (lastDate < range.startDate) {
        return { content: [{ type: "text", text: `The series starts ${range.startDate}, so nothing would be kept. Use delete_event with entireSeries instead.` }], isError: true };
      }
      await graph(token, `/me/events/${encodeURIComponent(masterId)}`, {
        method: "PATCH",
        body: { recurrence: { pattern, range: { type: "endDate", startDate: range.startDate, endDate: lastDate, recurrenceTimeZone: range.recurrenceTimeZone } } },
      });
      return ok(`"${master.subject}" now ends on ${lastDate}. Later dates were removed.`);
    }),
  );

  server.registerTool(
    "respond_to_event",
    {
      description: "Accept, decline, or tentatively accept a meeting invite.",
      inputSchema: z.object({
        id: z.string(),
        response: z.enum(["accept", "decline", "tentativelyAccept"]),
        message: z.string().optional(),
        notifyOrganizer: z.boolean().default(true),
      }),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    },
    safe(async ({ id, response, message, notifyOrganizer }) => {
      await graph(token, `/me/events/${encodeURIComponent(id)}/${response}`, {
        method: "POST",
        body: { comment: message ?? "", sendResponse: notifyOrganizer },
      });
      return ok(`Responded: ${response}.`);
    }),
  );

  return server;
}
