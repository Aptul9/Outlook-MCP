// Calendar reads over the same token and transport as mail. No writes, no RSVP.

import { q, type Obj, type Outlook } from "./api.ts";

export const MAX_RANGE_DAYS = 366;
export const MAX_EVENTS = 1000;
const EVENT_FIELDS = "Id,Subject,Start,End,IsAllDay,ShowAs,Location,Organizer,IsCancelled,Type,SeriesMasterId,OnlineMeetingUrl,WebLink";

export async function listCalendars(o: Outlook, account?: string): Promise<Obj[]> {
  const res = await o.request(account, "me/calendars?$top=100&$select=Id,Name,Color,IsDefaultCalendar,CanEdit,Owner");
  return (res.value ?? []) as Obj[];
}

/** An instant with an explicit zone (Z or +hh:mm); naive times are refused, the mailbox zone would be guessed. */
function instant(s: string, name: string): number {
  if (!/(Z|[+-]\d{2}:?\d{2})$/i.test(s.trim())) throw new Error(`${name} needs a time zone (Z or +hh:mm): ${s}`);
  const t = Date.parse(s);
  if (!Number.isFinite(t)) throw new Error(`${name} is not an ISO date-time: ${s}`);
  return t;
}

export async function listCalendarEvents(
  o: Outlook,
  account: string | undefined,
  start: string,
  end: string,
  opt: { calendar_id?: string; limit?: number; time_zone?: string } = {},
): Promise<{ events: Obj[]; has_more: boolean }> {
  const a = instant(start, "start");
  const b = instant(end, "end");
  if (b <= a) throw new Error("end must be after start");
  if (b - a > MAX_RANGE_DAYS * 86_400_000) throw new Error(`Calendar range exceeds ${MAX_RANGE_DAYS} days`);
  const limit = Math.max(1, Math.min(Math.trunc(opt.limit ?? 100) || 100, MAX_EVENTS));
  const base = opt.calendar_id ? `me/calendars/${q(opt.calendar_id)}/calendarview` : "me/calendarview";
  const qs = new URLSearchParams({ startDateTime: new Date(a).toISOString(), endDateTime: new Date(b).toISOString(), $top: String(limit), $select: EVENT_FIELDS, $orderby: "Start/DateTime" });
  const res = await o.request(account, `${base}?${qs}`, "GET", undefined, true, opt.time_zone ? { Prefer: `outlook.timezone="${opt.time_zone.replace(/"/g, "")}"` } : undefined);
  return { events: (res.value ?? []) as Obj[], has_more: !!res["@odata.nextLink"] };
}
