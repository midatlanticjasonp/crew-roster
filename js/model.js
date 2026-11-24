// Pure roster logic (no DOM): dates, layout of the Daily Breakdown, checks, new-week carry-over.

export const DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
export const DAY_NAMES = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"];

export const DEFAULT_CODES = [
  { code: "W", meaning: "Working / on board", label: "", manning: true, paid: true, noteRequired: false },
  { code: "ON", meaning: "Crew change – joined vessel (enter time)", label: "Joined", manning: true, paid: true, noteRequired: false },
  { code: "OFF", meaning: "Crew change – signed off (enter time)", label: "Signed off", manning: true, paid: true, noteRequired: false },
  { code: "T", meaning: "Travel / crew-change day", label: "Travel", manning: false, paid: true, noteRequired: false },
  { code: "TR", meaning: "Training / course", label: "Training", manning: false, paid: true, noteRequired: false },
  { code: "P", meaning: "Provisional – TBC, not confirmed", label: "TBC", manning: false, paid: false, noteRequired: false },
  { code: "S", meaning: "Sick", label: "Sick", manning: false, paid: true, noteRequired: true },
  { code: "A/L", meaning: "Annual leave (approved)", label: "Annual Leave", manning: false, paid: true, noteRequired: false },
  { code: "A", meaning: "Absent / not cleared to sail", label: "Absent", manning: false, paid: false, noteRequired: true },
  { code: "O", meaning: "Rest day / off", label: "Off", manning: false, paid: false, noteRequired: false },
];

// ---------- dates (all ISO 'YYYY-MM-DD', computed in UTC to avoid timezone drift) ----------
const toDate = (iso) => { const [y, m, d] = iso.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const toIso = (dt) => dt.toISOString().slice(0, 10);
export const addDays = (iso, n) => { const d = toDate(iso); d.setUTCDate(d.getUTCDate() + n); return toIso(d); };
export const isSunday = (iso) => toDate(iso).getUTCDay() === 0;
export const weekDates = (weekEnding) => DAYS.map((_, i) => addDays(weekEnding, i - 6));
export function sundayOnOrAfter(iso) { const d = toDate(iso); const add = (7 - d.getUTCDay()) % 7; return addDays(iso, add); }
export function fmtShort(iso) { if (!iso) return ""; const [y, m, d] = iso.split("-"); return `${d}.${m}.${y.slice(2)}`; }
export function fmtLong(iso) { const d = toDate(iso); return d.toLocaleDateString("en-GB", { weekday: "short", day: "2-digit", month: "short", year: "numeric", timeZone: "UTC" }); }
export function ordinal(n) { const s = ["th", "st", "nd", "rd"], v = n % 100; return n + (s[(v - 20) % 10] || s[v] || s[0]); }
export function dayHeader(iso, i) {
  const n = toDate(iso).getUTCDate();
  return `${DAY_NAMES[i]} ${String(n).padStart(2, "0")}${ordinal(n).replace(/^\d+/, "")}`;
}

// ---------- codes ----------
export const cssKey = (code) => (code || "").toUpperCase().replace(/[^A-Z]/g, "") || "X";
export const isCC = (code) => code === "ON" || code === "OFF";
export function codeMap(settings) { const m = new Map(); (settings.codes || DEFAULT_CODES).forEach((c) => m.set(c.code, c)); return m; }
export function entryText(code, time, cmap) {
  if (!code || code === "W") return "";
  if (code === "ON") return `JOINED${time ? " " + time : ""}`;
  if (code === "OFF") return `SIGNED OFF${time ? " " + time : ""}`;
  const c = cmap.get(code);
  return c ? (c.label || c.meaning || code) : code;
}
export const validTime = (t) => /^([01]\d|2[0-3]):[0-5]\d$/.test(t || "");

// ---------- helpers ----------
let _seq = 0;
export const uid = (p = "L") => p + Date.now().toString(36) + (++_seq).toString(36) + Math.random().toString(36).slice(2, 5);
export const emptyDays = () => DAYS.map(() => ({ code: "", time: "" }));
export function newLine(vessel, position) { return { id: uid(), vessel, position: position || "", crewId: "", days: emptyDays(), note: "", relievingCrewId: "" }; }
export function blankWeek(weekEnding, settings) {
  return { weekEnding, status: "draft", lines: [], portCaptain: settings.portCaptain || "", preparedBy: "", checkedBy: "", lookaheadPrintedAt: "", confirmedAt: "" };
}
export function positionOrder(vessel, settings) {
  const order = [];
  (vessel?.manning || []).forEach((m) => { if (!order.includes(m.position)) order.push(m.position); });
  (settings.positions || []).forEach((p) => { if (!order.includes(p)) order.push(p); });
  return order;
}
export function sortLines(lines, vessel, settings) {
  const order = positionOrder(vessel, settings);
  const rank = (p) => { const i = order.indexOf(p); return i < 0 ? 999 : i; };
  return lines.map((l, i) => [l, i]).sort((a, b) => rank(a[0].position) - rank(b[0].position) || a[1] - b[1]).map((x) => x[0]);
}
export const hasDays = (l) => l.days.some((d) => d.code);
export function lineStats(line, cmap) {
  let onBoard = 0, paid = 0;
  line.days.forEach((d) => { if (!d.code) return; const c = cmap.get(d.code); if (c?.manning) onBoard++; if (c?.paid) paid++; });
  return { onBoard, paid };
}

// ---------- Daily Breakdown layout ----------
// Every line with at least one day filled gets printed. A line whose "relieving" person is on the same
// vessel + position shares that person's printed line (crew change in one cell).
export function buildBreakdown(week, meta, crewById, opts = {}) {
  const cmap = codeMap(meta.settings);
  const out = [];
  for (const vessel of meta.vessels) {
    const vLines = week.lines.filter((l) => l.vessel === vessel.id);
    const manningReq = new Map((vessel.manning || []).map((m) => [m.position, m.required || 0]));
    if (!vLines.length && !opts.includeEmptyVessels && ![...manningReq.values()].some((x) => x > 0)) continue;
    const order = positionOrder(vessel, meta.settings);
    const positions = [...new Set([...order.filter((p) => manningReq.has(p) || vLines.some((l) => l.position === p)), ...vLines.map((l) => l.position)])];
    const groups = [];
    const notes = [];
    for (const pos of positions) {
      const gl = vLines.filter((l) => l.position === pos && l.crewId && hasDays(l));
      // assign slots
      const slotOf = new Map(); const slots = [];
      const byCrew = new Map(gl.map((l) => [l.crewId, l]));
      for (const l of gl) {
        const rel = l.relievingCrewId && byCrew.get(l.relievingCrewId);
        if (!rel || rel === l) { slotOf.set(l.id, slots.length); slots.push([l]); }
      }
      for (let pass = 0; pass < 5; pass++) {
        for (const l of gl) {
          if (slotOf.has(l.id)) continue;
          const rel = byCrew.get(l.relievingCrewId);
          if (rel && slotOf.has(rel.id)) { const s = slotOf.get(rel.id); slotOf.set(l.id, s); slots[s].push(l); }
        }
      }
      for (const l of gl) if (!slotOf.has(l.id)) { slotOf.set(l.id, slots.length); slots.push([l]); } // circular relieving → own line
      const rows = slots.map((sl) => DAYS.map((_, di) => {
        const ents = sl.filter((l) => l.days[di].code).map((l) => {
          const d = l.days[di];
          return { name: crewById.get(l.crewId)?.name || "?", code: d.code, time: d.time, text: entryText(d.code, d.time, cmap), key: isCC(d.code) ? "CC" : cssKey(d.code) };
        });
        ents.sort((a, b) => (a.code === "OFF" ? -1 : b.code === "OFF" ? 1 : a.code === "ON" ? 1 : b.code === "ON" ? -1 : 0));
        return ents;
      }));
      const req = manningReq.get(pos) || 0;
      if (opts.showEmpty !== false) for (let i = rows.length; i < req; i++) rows.push(DAYS.map(() => []));
      if (rows.length) groups.push({ position: pos, rows, required: req });
      gl.forEach((l) => { if (l.note && l.note.trim()) notes.push({ name: crewById.get(l.crewId)?.name || "?", note: l.note.trim() }); });
    }
    // manning gaps (screen + checks)
    const gaps = [];
    const offHire = /OFF[\s-]?HIRE|DRY[\s-]?DOCK|LAID[\s-]?UP/i.test(vessel.status || "");
    if (!offHire) for (const [pos, req] of manningReq) {
      if (!req) continue;
      const short = [];
      DAYS.forEach((_, di) => {
        const n = vLines.filter((l) => l.position === pos && l.crewId && cmap.get(l.days[di].code)?.manning).length;
        if (n < req) short.push({ day: di, have: n, req });
      });
      if (short.length) gaps.push({ position: pos, short });
    }
    out.push({ vessel, groups, notes, gaps });
  }
  return out;
}

// ---------- checks ----------
export function runChecks(week, meta, crewById) {
  const cmap = codeMap(meta.settings);
  const vById = new Map(meta.vessels.map((v) => [v.id, v]));
  const issues = [];
  const add = (sev, line, msg, cat) => issues.push({ sev, lineId: line?.id || "", vesselId: line?.vessel || "", msg, cat });
  const nm = (l) => crewById.get(l.crewId)?.name || "(no name)";
  for (const l of week.lines) {
    const filled = hasDays(l);
    const v = vById.get(l.vessel);
    if (!v) add("block", l, `${nm(l)}: vessel no longer exists – move the line to a vessel`, "Lines");
    if (!l.crewId) { if (filled) add("block", l, `Days entered with no crew name (${v?.name || "?"} – ${l.position || "no position"})`, "Lines"); continue; }
    if (!crewById.get(l.crewId)) add("block", l, `Crew member not found in the crew list`, "Lines");
    if (!l.position) add("block", l, `${nm(l)}: no position selected`, "Lines");
    if (!filled) add("warn", l, `${nm(l)} (${v?.name || "?"}): no days filled – will not print`, "Lines");
    l.days.forEach((d, i) => {
      if (!d.code) return;
      if (!cmap.has(d.code)) add("block", l, `${nm(l)}: "${d.code}" on ${DAYS[i]} is not a recognised code`, "Codes");
      if (isCC(d.code) && !validTime(d.time)) add("block", l, `${nm(l)}: ${d.code === "ON" ? "joining" : "sign-off"} time missing on ${DAYS[i]}`, "Crew changes");
    });
    const needNote = [...new Set(l.days.map((d) => d.code).filter((c) => cmap.get(c)?.noteRequired))];
    if (needNote.length && !(l.note || "").trim()) add("block", l, `${nm(l)}: ${needNote.map((c) => cmap.get(c).label || c).join(" / ")} needs a note for payroll`, "Notes");
    if (l.days.some((d) => d.code === "P")) add("block", l, `${nm(l)}: provisional (TBC) days – confirm before the confirmed print`, "Provisional");
    if (l.relievingCrewId) {
      const rel = week.lines.find((x) => x.crewId === l.relievingCrewId && x.vessel === l.vessel && x.position === l.position);
      if (!rel) add("warn", l, `${nm(l)} is set as relieving ${crewById.get(l.relievingCrewId)?.name || "?"}, who is not on ${v?.name || "this vessel"} as ${l.position} – printed on own line`, "Crew changes");
    }
  }
  // double booking: same person on duty twice the same day (an OFF on one line + ON on another is fine)
  const byCrew = new Map();
  week.lines.forEach((l) => { if (l.crewId) (byCrew.get(l.crewId) || byCrew.set(l.crewId, []).get(l.crewId)).push(l); });
  for (const [cid, ls] of byCrew) {
    if (ls.length < 2) continue;
    DAYS.forEach((_, i) => {
      const codes = ls.map((l) => l.days[i].code).filter(Boolean);
      const duty = codes.filter((c) => c === "W" || c === "P" || isCC(c));
      const cc = codes.filter(isCC);
      const ok = duty.length <= 1 || (duty.length === 2 && cc.length === 2 && cc.includes("ON") && cc.includes("OFF"));
      if (!ok) add("block", ls[0], `${crewById.get(cid)?.name}: on duty on two lines on ${DAYS[i]}`, "Double booking");
    });
  }
  // manning gaps
  const bd = buildBreakdown(week, meta, crewById, { includeEmptyVessels: true });
  for (const b of bd) for (const g of b.gaps) {
    const days = g.short.map((s) => DAYS[s.day]).join(", ");
    issues.push({ sev: "warn", lineId: "", vesselId: b.vessel.id, msg: `${b.vessel.name} – ${g.position}: short-handed on ${days} (${g.short[0].req} required)`, cat: "Manning" });
  }
  // certificates
  const end = week.weekEnding;
  const seen = new Set();
  for (const l of week.lines) {
    const c = crewById.get(l.crewId); if (!c || seen.has(c.id) || !l.days.some((d) => cmap.get(d.code)?.manning)) continue;
    seen.add(c.id);
    const st = certStatus(c, end);
    if (st.cls === "exp") add("warn", l, `${c.name}: on duty with an expired certificate (${st.text})`, "Certificates");
  }
  if (!(week.portCaptain || "").trim()) issues.push({ sev: "block", lineId: "", vesselId: "", msg: "Port Captain name missing (top of the Roster tab)", cat: "Sign-off" });
  return issues;
}

export function certStatus(c, weekEnding) {
  const dates = [["CoC", c.cocExpiry], ["Medical", c.medicalExpiry], ["Offshore", c.offshoreExpiry]].filter((x) => x[1]);
  if (!dates.length) return { cls: "none", text: "No dates" };
  dates.sort((a, b) => (a[1] < b[1] ? -1 : 1));
  const [what, d] = dates[0];
  const weekStart = addDays(weekEnding, -6);
  if (d < weekStart) return { cls: "exp", text: `${what} expired ${fmtShort(d)}` };
  if (d <= addDays(weekEnding, 30)) return { cls: "due", text: `${what} due ${fmtShort(d)}` };
  return { cls: "ok", text: "OK" };
}

// ---------- new week ----------
// Carry everyone whose last entry is not OFF; prefill their last code for all 7 days.
export function startNewWeek(prev, settings, crewById) {
  const next = blankWeek(addDays(prev.weekEnding, 7), settings);
  next.portCaptain = prev.portCaptain || settings.portCaptain || "";
  next.preparedBy = prev.preparedBy || "";
  const carried = [], removed = [];
  for (const l of prev.lines) {
    const name = crewById.get(l.crewId)?.name || "(no name)";
    let last = "";
    for (let i = 6; i >= 0; i--) if (l.days[i].code) { last = l.days[i].code; break; }
    if (!l.crewId || !last) { if (l.crewId) removed.push({ name, why: "no days entered" }); continue; }
    if (last === "OFF") { removed.push({ name, why: "signed off" }); continue; }
    if (crewById.get(l.crewId)?.deleted) { removed.push({ name, why: "deleted from crew list" }); continue; }
    const fill = last === "ON" || last === "T" ? "W" : last;
    next.lines.push({ id: uid(), vessel: l.vessel, position: l.position, crewId: l.crewId, days: DAYS.map(() => ({ code: fill, time: "" })),
      note: fill === "W" ? "" : (l.note || ""), relievingCrewId: "" });
    carried.push({ name, fill });
  }
  return { week: next, carried, removed };
}
