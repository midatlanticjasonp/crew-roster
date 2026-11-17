import { createStore } from "./store.js";
import * as M from "./model.js";
import { exportExcel } from "./export.js";

const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

const DEFAULT_POSITIONS = ["CAPTAIN", "MATE", "MATE / DPO", "MATE - TRAINEE", "CHIEF ENG", "ENGINEER 2", "OILER / SAILOR - OS", "COOK", "SAILOR - AB", "SAILOR - OS"];

const S = {
  store: null, mode: "local", user: null,
  meta: null, crew: [], crewById: new Map(), week: null,
  printStamp: null, tab: "roster",
};

// ---------------------------------------------------------------- saving
const timers = {};
function scheduleSave(kind) {
  setSave("Saving…");
  clearTimeout(timers[kind]);
  timers[kind] = setTimeout(async () => {
    try {
      if (kind === "week") await S.store.saveWeek(S.week);
      if (kind === "meta") await S.store.saveMeta(S.meta);
      if (kind === "crew") await S.store.saveCrew(S.crew);
      setSave("All changes saved");
    } catch (e) { console.error(e); setSave("⚠ Save failed – check connection"); }
  }, 500);
}
function setSave(t) { $("#saveState").textContent = t; }
async function flushSaves() {
  for (const k of Object.keys(timers)) clearTimeout(timers[k]);
  await S.store.saveWeek(S.week); await S.store.saveMeta(S.meta); await S.store.saveCrew(S.crew);
  setSave("All changes saved");
}

function weekChanged({ full = false } = {}) {
  if (S.week.status === "confirmed") { S.week.status = "amended"; renderStatus(); }
  scheduleSave("week");
  if (full) renderRoster();
  refreshChecksBadge();
  if (S.tab === "breakdown") renderBreakdown();
  if (S.tab === "checks") renderChecks();
}
function reindexCrew() { S.crewById = new Map(S.crew.map((c) => [c.id, c])); }
const crewName = (id) => S.crewById.get(id)?.name || "";
const settings = () => S.meta.settings;
const cmap = () => M.codeMap(settings());

// ---------------------------------------------------------------- dialogs
function dialog(title, html, { ok = "OK", cancel = "Cancel", danger = false } = {}) {
  const d = $("#dlg");
  $("#dlgTitle").textContent = title; $("#dlgBody").innerHTML = html;
  $("#dlgOk").textContent = ok; $("#dlgOk").className = "btn " + (danger ? "danger" : "primary");
  $("#dlgCancel").textContent = cancel; $("#dlgCancel").hidden = !cancel;
  d.showModal();
  return new Promise((res) => d.addEventListener("close", () => res(d.returnValue === "ok"), { once: true }));
}

// ---------------------------------------------------------------- header / week navigation
function renderHeader() {
  $("#companyName").textContent = settings().company || "Set company name in Settings";
  const ids = [...S.meta.weekIds].sort().reverse();
  $("#weekSelect").innerHTML = ids.map((id) => `<option value="${id}" ${id === S.week.weekEnding ? "selected" : ""}>${M.fmtShort(id)}  (${M.fmtShort(M.addDays(id, -6))} – ${M.fmtShort(id)})</option>`).join("");
  const sorted = [...S.meta.weekIds].sort(); const i = sorted.indexOf(S.week.weekEnding);
  $("#prevWeek").disabled = i <= 0; $("#nextWeek").disabled = i < 0 || i >= sorted.length - 1;
  $("#wkPortCaptain").value = S.week.portCaptain || "";
  $("#wkPrepared").value = S.week.preparedBy || "";
  $("#wkChecked").value = S.week.checkedBy || "";
  $("#legend").innerHTML = (settings().codes || []).map((c) => `<span class="c-${M.cssKey(c.code)}" title="${esc(c.meaning)}"><b>${esc(c.code)}</b> ${esc(c.label || c.meaning)}</span>`).join("")
    + `<span title="Choosing a code on an otherwise empty line fills the rest of the week. ON fills W after it; OFF clears the days after it.">ⓘ quick-fill</span>`;
  renderStatus();
}
function renderStatus() {
  const st = S.week.status || "draft";
  const p = $("#statusPill");
  const txt = { draft: "Draft", lookahead: "Look-ahead sent", confirmed: "Confirmed", amended: "Changed after confirm" }[st] || st;
  p.textContent = txt; p.className = "status-pill " + st;
  p.title = [S.week.lookaheadPrintedAt && "Look-ahead printed " + new Date(S.week.lookaheadPrintedAt).toLocaleString(), S.week.confirmedAt && "Confirmed " + new Date(S.week.confirmedAt).toLocaleString()].filter(Boolean).join("\n");
}
async function openWeek(id) {
  if (S.week) await S.store.saveWeek(S.week);
  const w = await S.store.loadWeek(id);
  S.week = w || M.blankWeek(id, settings());
  normaliseWeek(S.week);
  renderAll();
}
function normaliseWeek(w) {
  w.lines = (w.lines || []).map((l) => ({ ...M.newLine(l.vessel, l.position), ...l, days: M.DAYS.map((_, i) => ({ code: l.days?.[i]?.code || "", time: l.days?.[i]?.time || "" })) }));
}

// ---------------------------------------------------------------- ROSTER
function codeOptions(sel) {
  return `<option value=""></option>` + (settings().codes || []).map((c) => `<option value="${esc(c.code)}" ${c.code === sel ? "selected" : ""}>${esc(c.code)}</option>`).join("")
    + (sel && !cmap().has(sel) ? `<option value="${esc(sel)}" selected>${esc(sel)}?</option>` : "");
}
function positionOptions(sel) {
  const list = [...(settings().positions || [])]; if (sel && !list.includes(sel)) list.push(sel);
  return `<option value="">— position —</option>` + list.map((p) => `<option ${p === sel ? "selected" : ""}>${esc(p)}</option>`).join("");
}
function dayCell(l, i) {
  const d = l.days[i];
  const cc = M.isCC(d.code);
  const bad = (cc && !M.validTime(d.time)) || (d.code && !cmap().has(d.code));
  return `<td class="day c-${d.code ? M.cssKey(d.code) : "X"} ${bad ? "err" : ""}" data-day="${i}">
    <div class="daycell"><select data-f="code" data-day="${i}" aria-label="${M.DAYS[i]} code">${codeOptions(d.code)}</select>${cc ? `<input type="time" data-f="time" data-day="${i}" value="${esc(d.time)}" aria-label="${M.DAYS[i]} time" required>` : ""}</div></td>`;
}
function lineRow(l, issuesByLine) {
  const st = M.lineStats(l, cmap());
  const iss = issuesByLine.get(l.id) || [];
  const block = iss.filter((x) => x.sev === "block");
  const others = S.week.lines.filter((x) => x.vessel === l.vessel && x.position === l.position && x.crewId && x.crewId !== l.crewId);
  const relOpts = `<option value="">—</option>` + others.map((x) => `<option value="${x.crewId}" ${x.crewId === l.relievingCrewId ? "selected" : ""}>${esc(crewName(x.crewId))}</option>`).join("")
    + (l.relievingCrewId && !others.some((x) => x.crewId === l.relievingCrewId) ? `<option value="${l.relievingCrewId}" selected>${esc(crewName(l.relievingCrewId))} (not here)</option>` : "");
  return `<tr data-line="${l.id}" class="${block.length ? "flagged" : ""}">
    <td class="flag" title="${esc(iss.map((x) => x.msg).join("\n"))}">${block.length ? "!" : iss.length ? "·" : ""}</td>
    <td style="min-width:165px"><select data-f="position" aria-label="Position">${positionOptions(l.position)}</select></td>
    <td class="name ${l.crewId && !S.crewById.get(l.crewId) ? "err" : ""}"><input data-f="crew" list="crewNames" value="${esc(crewName(l.crewId))}" placeholder="Type or pick a name" aria-label="Crew name"></td>
    <td><select data-f="relieving" aria-label="Relieving" title="Crew change: who this person replaces. They then share one line on the printout.">${relOpts}</select></td>
    ${M.DAYS.map((_, i) => dayCell(l, i)).join("")}
    <td class="ro">${st.onBoard}</td><td class="ro">${st.paid}</td>
    <td class="note"><input data-f="note" value="${esc(l.note)}" placeholder="Note for payroll" aria-label="Note"></td>
    <td style="width:28px"><button class="rowbtn" data-act="del" title="Remove line" aria-label="Remove line">×</button></td>
  </tr>`;
}
function renderRoster() {
  const dates = M.weekDates(S.week.weekEnding);
  const issues = M.runChecks(S.week, S.meta, S.crewById);
  const issuesByLine = new Map(); issues.forEach((x) => { if (x.lineId) (issuesByLine.get(x.lineId) || issuesByLine.set(x.lineId, []).get(x.lineId)).push(x); });
  let html = `<thead><tr><th></th><th>Position</th><th>Crew</th><th title="Who they replace (crew change)">Relieving</th>
    ${dates.map((d, i) => `<th class="wk">${M.DAYS[i]}<small>${M.fmtShort(d).slice(0, 5)}</small></th>`).join("")}
    <th title="Days counting toward manning">On board</th><th>Paid</th><th>Note</th><th></th></tr></thead><tbody>`;
  const vesselIds = new Set(S.meta.vessels.map((v) => v.id));
  const groups = [...S.meta.vessels];
  const orphans = S.week.lines.filter((l) => !vesselIds.has(l.vessel));
  if (orphans.length) groups.push({ id: "__orphan", name: "Lines with a deleted vessel", manning: [] });
  for (const v of groups) {
    const lines = v.id === "__orphan" ? orphans : M.sortLines(S.week.lines.filter((l) => l.vessel === v.id), v, settings());
    html += `<tr class="vessel-row"><td colspan="15">${esc(v.name)}<span class="vmeta">${esc([v.client && "with " + v.client, v.status].filter(Boolean).join(" · "))} · ${lines.filter((l) => l.crewId).length} crew lines</span>
      ${v.id !== "__orphan" ? `<button class="btn small" data-act="add" data-vessel="${v.id}">+ Add line</button>` : ""}</td></tr>`;
    html += lines.map((l) => lineRow(l, issuesByLine)).join("");
  }
  if (!S.meta.vessels.length) html += `<tr><td colspan="15" style="padding:16px">No vessels yet. Add them on the <b>Vessels &amp; Manning</b> tab, or load your starter data in <b>Settings &amp; Data</b>.</td></tr>`;
  html += "</tbody>";
  $("#rosterTable").innerHTML = html;
  $("#crewNames").innerHTML = S.crew.filter((c) => c.active !== false).map((c) => `<option value="${esc(c.name)}">${esc([c.fullName, c.rank].filter(Boolean).join(" · "))}</option>`).join("");
}
function lineById(id) { return S.week.lines.find((l) => l.id === id); }
function refreshRow(l) {
  const tr = $(`tr[data-line="${l.id}"]`); if (!tr) return;
  const issues = M.runChecks(S.week, S.meta, S.crewById);
  const by = new Map(); issues.forEach((x) => { if (x.lineId) (by.get(x.lineId) || by.set(x.lineId, []).get(x.lineId)).push(x); });
  const focusDay = document.activeElement?.closest("tr") === tr ? document.activeElement.dataset : null;
  tr.outerHTML = lineRow(l, by);
  if (focusDay) { const el = $(`tr[data-line="${l.id}"] [data-f="${focusDay.f}"]${focusDay.day !== undefined ? `[data-day="${focusDay.day}"]` : ""}`); el?.focus(); }
}

async function resolveCrewName(raw, position) {
  const name = raw.trim().toUpperCase().replace(/\s+/g, " ").replace(/^([A-Z])\s+(?=[A-Z])/, "$1. ");
  if (!name) return "";
  const hit = S.crew.find((c) => c.name === name) || S.crew.find((c) => c.name.replace(/[.\s]/g, "") === name.replace(/[.\s]/g, ""));
  if (hit) return hit.id;
  const ok = await dialog("New crew member?", `<p><b>${esc(name)}</b> isn't in the crew list.</p><p>Add them now? You can fill in their full name, phone and certificates on the Crew tab later.</p>`, { ok: "Add to crew list" });
  if (!ok) return null;
  const c = { id: nextCrewId(), name, fullName: "", rank: position || "", phone: "", email: "", dayRate: null, cocExpiry: "", medicalExpiry: "", offshoreExpiry: "", active: true, notes: "" };
  S.crew.push(c); reindexCrew(); scheduleSave("crew");
  return c.id;
}
function nextCrewId() {
  const n = Math.max(0, ...S.crew.map((c) => +(String(c.id).match(/(\d+)$/)?.[1] || 0))) + 1;
  return "NS-" + String(n).padStart(3, "0");
}

function bindRoster() {
  const t = $("#rosterTable");
  t.addEventListener("change", async (e) => {
    const el = e.target; const tr = el.closest("tr[data-line]"); if (!tr) return;
    const l = lineById(tr.dataset.line); if (!l) return;
    const f = el.dataset.f;
    if (f === "code") {
      const i = +el.dataset.day; const code = el.value;
      const othersEmpty = l.days.every((d, j) => j === i || !d.code);
      l.days[i] = { code, time: M.isCC(code) ? l.days[i].time : "" };
      if (code === "ON") for (let j = i + 1; j < 7; j++) { if (!l.days[j].code) l.days[j] = { code: "W", time: "" }; }
      else if (code === "OFF") for (let j = i + 1; j < 7; j++) l.days[j] = { code: "", time: "" };
      else if (code && othersEmpty) for (let j = i + 1; j < 7; j++) l.days[j] = { code, time: "" };
      refreshRow(l);
      if (M.isCC(code)) $(`tr[data-line="${l.id}"] input[data-f="time"][data-day="${i}"]`)?.focus();
      weekChanged();
    } else if (f === "time") {
      l.days[+el.dataset.day].time = el.value; refreshRow(l); weekChanged();
    } else if (f === "position") {
      l.position = el.value; weekChanged({ full: true });
    } else if (f === "relieving") {
      l.relievingCrewId = el.value; weekChanged();
    } else if (f === "crew") {
      const id = await resolveCrewName(el.value, l.position);
      if (id === null) { el.value = crewName(l.crewId); return; }
      l.crewId = id; weekChanged({ full: true });
    }
  });
  t.addEventListener("input", (e) => {
    if (e.target.dataset.f !== "note") return;
    const l = lineById(e.target.closest("tr").dataset.line); l.note = e.target.value; weekChanged();
  });
  t.addEventListener("click", async (e) => {
    const b = e.target.closest("[data-act]"); if (!b) return;
    if (b.dataset.act === "add") {
      const v = S.meta.vessels.find((x) => x.id === b.dataset.vessel);
      const l = M.newLine(v.id, ""); S.week.lines.push(l); weekChanged({ full: true });
      $(`tr[data-line="${l.id}"] select[data-f="position"]`)?.focus();
    }
    if (b.dataset.act === "del") {
      const l = lineById(b.closest("tr").dataset.line);
      if (l.crewId || M.hasDays(l)) { if (!(await dialog("Remove line?", `<p>Remove <b>${esc(crewName(l.crewId) || "this line")}</b> from this week's roster?</p>`, { ok: "Remove", danger: true }))) return; }
      S.week.lines = S.week.lines.filter((x) => x !== l); weekChanged({ full: true });
    }
  });
  for (const [id, key] of [["#wkPortCaptain", "portCaptain"], ["#wkPrepared", "preparedBy"], ["#wkChecked", "checkedBy"]]) {
    $(id).addEventListener("input", (e) => { S.week[key] = e.target.value; weekChanged(); });
  }
}

// ---------------------------------------------------------------- DAILY BREAKDOWN
function stampFor() {
  if (S.printStamp) return S.printStamp;
  const st = S.week.status;
  if (st === "confirmed") return { mode: "confirmed", at: S.week.confirmedAt };
  if (st === "lookahead") return { mode: "lookahead", at: S.week.lookaheadPrintedAt };
  return { mode: "draft", at: "" };
}
function renderBreakdown() {
  const dates = M.weekDates(S.week.weekEnding);
  const bd = M.buildBreakdown(S.week, S.meta, S.crewById, { showEmpty: $("#showEmptyLines").checked });
  const stamp = stampFor();
  const stampTxt = { draft: "DRAFT – not sent", lookahead: "LOOK-AHEAD", confirmed: "CONFIRMED" }[stamp.mode]
    + (stamp.at ? " · printed " + new Date(stamp.at).toLocaleString("en-GB", { day: "2-digit", month: "2-digit", year: "2-digit", hour: "2-digit", minute: "2-digit" }) : "");
  const pages = bd.map(({ vessel, groups, notes, gaps }) => {
    const totalRows = groups.reduce((n, g) => n + g.rows.length, 0) || 1;
    let rows = ""; let first = true;
    const vcell = `<td class="vcell" rowspan="${totalRows}"><div class="vname">${esc(vessel.name.replace(/^MV\s+/i, ""))}</div>
      ${vessel.client || vessel.status ? `<div>${esc((vessel.status || "IN SERVICE") + (vessel.client ? " WITH " + vessel.client : ""))}</div>` : ""}
      ${vessel.contact ? `<div>${esc(vessel.contact)}</div>` : ""}</td>`;
    for (const g of groups) {
      g.rows.forEach((r, ri) => {
        rows += "<tr>" + (first ? vcell : "") + (ri === 0 ? `<td class="pos" rowspan="${g.rows.length}">${esc(g.position)}</td>` : "");
        first = false;
        rows += r.map((ents) => {
          if (!ents.length) return `<td class="${ri >= 0 && g.rows.indexOf(r) >= 0 && r.every((x) => !x.length) ? "" : ""}"></td>`;
          const cls = ents.length === 1 ? (ents[0].key === "W" ? "" : "k-" + ents[0].key) : (ents.some((e) => e.key === "CC") ? "k-CC" : "");
          return `<td class="${cls}">${ents.map((e) => `<span class="ent ${ents.length > 1 && e.key !== "CC" && e.key !== "W" ? "k-" + e.key : ""}">${esc(e.name)}${e.text ? `<span class="lbl">${esc(e.text)}</span>` : ""}</span>`).join("")}</td>`;
        }).join("") + "</tr>";
      });
    }
    if (!groups.length) rows = `<tr>${vcell}<td class="pos"></td>${dates.map(() => "<td></td>").join("")}</tr>`;
    const gapTxt = gaps.length ? `<div class="screen-gap">Short-handed (not printed): ${gaps.map((g) => `${esc(g.position)} – ${g.short.map((s) => M.DAYS[s.day]).join(", ")}`).join(" · ")}</div>` : "";
    return `<section class="vpage">
      <div class="vp-title">VESSELS - WEEKLY ROSTER <span class="vp-stamp ${stamp.mode}">${esc(stampTxt)}</span></div>
      <div class="vp-head"><div>COMPANY:<span>${esc(settings().company)}</span></div><div>Week Ending:<span>${M.fmtShort(S.week.weekEnding)}</span></div><div>Port Captain:<span>${esc(S.week.portCaptain)}</span></div></div>
      <table class="bd"><colgroup><col style="width:12%"><col style="width:13%">${dates.map(() => `<col style="width:${75 / 7}%">`).join("")}</colgroup>
        <thead><tr><th>VESSEL</th><th>CREW</th>${dates.map((d, i) => `<th>${M.dayHeader(d, i)}</th>`).join("")}</tr></thead>
        <tbody>${rows}</tbody></table>
      <div class="vp-notes"><b>Notes:</b><ul>${notes.map((n) => `<li>${esc(n.name)} – ${esc(n.note)}</li>`).join("") || "<li>&nbsp;</li>"}</ul></div>
      <div class="vp-sign"><div>Prepared by:<span>${esc(S.week.preparedBy)}</span></div><div>Checked by:<span>${esc(S.week.checkedBy)}</span></div><div>Port Captain:<span></span></div></div>
      ${gapTxt}
    </section>`;
  });
  $("#breakdown").innerHTML = pages.join("") || `<p class="hint">Nothing to show yet.</p>`;
}

// ---------------------------------------------------------------- CHECKS
function refreshChecksBadge() {
  const issues = M.runChecks(S.week, S.meta, S.crewById);
  const blocks = issues.filter((x) => x.sev === "block").length;
  const b = $("#checksBadge");
  b.hidden = !issues.length; b.textContent = blocks || issues.length; b.className = "badge" + (blocks ? "" : " warn");
  b.title = `${blocks} to fix before the confirmed print, ${issues.length - blocks} warnings`;
}
function renderChecks() {
  const issues = M.runChecks(S.week, S.meta, S.crewById);
  if (!issues.length) { $("#checksList").innerHTML = `<div class="allgood">✓ No problems found – ready for the confirmed print.</div>`; return; }
  const cats = [...new Set(issues.map((x) => x.cat))];
  const blocks = issues.filter((x) => x.sev === "block").length;
  $("#checksList").innerHTML = `<p class="hint">${blocks ? `<b>${blocks}</b> to fix before the confirmed print` : "Nothing blocking"} · ${issues.length - blocks} warnings.</p>` +
    cats.map((c) => {
      const list = issues.filter((x) => x.cat === c);
      return `<div class="check-group"><h3>${esc(c)} <span class="badge ${list.some((x) => x.sev === "block") ? "" : "warn"}">${list.length}</span></h3><ul>${list.map((x) =>
        `<li><span class="sev ${x.sev}">${x.sev === "block" ? "FIX" : "CHECK"}</span><span>${esc(x.msg)}</span>${x.lineId ? `<button class="btn small" data-jump="${x.lineId}">Go to line</button>` : x.cat === "Sign-off" ? `<button class="btn small" data-jump="__pc">Go to</button>` : ""}</li>`).join("")}</ul></div>`;
    }).join("");
}
function jumpTo(lineId) {
  switchTab("roster");
  if (lineId === "__pc") { $("#wkPortCaptain").focus(); return; }
  const tr = $(`tr[data-line="${lineId}"]`); if (!tr) return;
  tr.scrollIntoView({ block: "center" }); tr.classList.add("jump"); setTimeout(() => tr.classList.remove("jump"), 1700);
  tr.querySelector("input,select")?.focus({ preventScroll: true });
}

// ---------------------------------------------------------------- CREW
function renderCrew() {
  const q = $("#crewSearch").value.trim().toUpperCase();
  const showInactive = $("#crewShowInactive").checked;
  const onWeek = new Map(); S.week.lines.forEach((l) => { if (l.crewId && M.hasDays(l)) onWeek.set(l.crewId, S.meta.vessels.find((v) => v.id === l.vessel)?.name || ""); });
  const list = S.crew.filter((c) => (showInactive || c.active !== false) && (!q || [c.name, c.fullName, c.rank, c.phone, c.id].join(" ").toUpperCase().includes(q)));
  const posOpts = (sel) => `<option value=""></option>` + [...new Set([...(settings().positions || []), sel].filter(Boolean))].map((p) => `<option ${p === sel ? "selected" : ""}>${esc(p)}</option>`).join("");
  $("#crewTable").innerHTML = `<thead><tr><th>ID</th><th>Roster name</th><th>Full name</th><th>Rank</th><th>Phone</th><th>Email</th><th>CoC expiry</th><th>Medical expiry</th><th>Offshore safety expiry</th><th>Certificates</th><th>This week</th><th>Active</th><th>Notes</th></tr></thead><tbody>` +
    list.map((c) => { const cs = M.certStatus(c, S.week.weekEnding); return `<tr data-crew="${esc(c.id)}">
      <td><input data-cf="id" value="${esc(c.id)}" style="width:80px"></td>
      <td><input data-cf="name" value="${esc(c.name)}" style="min-width:150px;font-weight:600"></td>
      <td><input data-cf="fullName" value="${esc(c.fullName)}" style="min-width:150px"></td>
      <td style="min-width:170px"><select data-cf="rank">${posOpts(c.rank)}</select></td>
      <td><input data-cf="phone" value="${esc(c.phone)}" style="width:110px"></td>
      <td><input data-cf="email" value="${esc(c.email)}" type="email" style="min-width:170px"></td>
      <td><input data-cf="cocExpiry" type="date" value="${esc(c.cocExpiry)}"></td>
      <td><input data-cf="medicalExpiry" type="date" value="${esc(c.medicalExpiry)}"></td>
      <td><input data-cf="offshoreExpiry" type="date" value="${esc(c.offshoreExpiry)}"></td>
      <td class="cert ${cs.cls}">${esc(cs.text)}</td>
      <td class="ro">${esc((onWeek.get(c.id) || "").replace(/^MV /, ""))}</td>
      <td style="text-align:center"><input data-cf="active" type="checkbox" ${c.active !== false ? "checked" : ""} style="width:auto;min-height:0"></td>
      <td><input data-cf="notes" value="${esc(c.notes)}" style="min-width:220px"></td></tr>`; }).join("") + "</tbody>";
}
function bindCrew() {
  $("#crewSearch").addEventListener("input", renderCrew);
  $("#crewShowInactive").addEventListener("change", renderCrew);
  $("#addCrew").addEventListener("click", () => {
    const c = { id: nextCrewId(), name: "NEW CREW", fullName: "", rank: "", phone: "", email: "", dayRate: null, cocExpiry: "", medicalExpiry: "", offshoreExpiry: "", active: true, notes: "" };
    S.crew.unshift(c); reindexCrew(); scheduleSave("crew"); $("#crewSearch").value = ""; renderCrew();
    const el = $(`tr[data-crew="${c.id}"] [data-cf="name"]`); el.focus(); el.select();
  });
  $("#crewTable").addEventListener("change", async (e) => {
    const el = e.target; const tr = el.closest("tr[data-crew]"); if (!tr) return;
    const c = S.crew.find((x) => x.id === tr.dataset.crew); const f = el.dataset.cf;
    if (f === "active") c.active = el.checked;
    else if (f === "name") {
      const v = el.value.trim().toUpperCase().replace(/\s+/g, " ");
      if (S.crew.some((x) => x !== c && x.name === v)) { await dialog("Name already used", `<p><b>${esc(v)}</b> is already in the crew list. Roster names must be unique so payroll can't mix people up – add a first name or initial.</p>`, { cancel: "" }); el.value = c.name; return; }
      c.name = v; el.value = v;
    } else if (f === "id") {
      const v = el.value.trim(); if (!v || S.crew.some((x) => x !== c && x.id === v)) { el.value = c.id; return; }
      S.week.lines.forEach((l) => { if (l.crewId === c.id) l.crewId = v; if (l.relievingCrewId === c.id) l.relievingCrewId = v; });
      c.id = v; tr.dataset.crew = v; scheduleSave("week");
    } else c[f] = el.value;
    reindexCrew(); scheduleSave("crew");
    if (["cocExpiry", "medicalExpiry", "offshoreExpiry"].includes(f)) renderCrew();
    refreshChecksBadge();
  });
}

// ---------------------------------------------------------------- VESSELS
function renderVessels() {
  const posOpts = (sel) => [...new Set([...(settings().positions || []), sel].filter(Boolean))].map((p) => `<option ${p === sel ? "selected" : ""}>${esc(p)}</option>`).join("");
  $("#vesselsList").innerHTML = S.meta.vessels.map((v, vi) => `<div class="card" data-vessel="${v.id}">
    <h3>${esc(v.name || "New vessel")}</h3>
    <div class="grid2">
      <div class="field"><label>Vessel name</label><input data-vf="name" value="${esc(v.name)}"></div>
      <div class="field"><label>Client / charter</label><input data-vf="client" value="${esc(v.client)}"></div>
      <div class="field"><label>Vessel contact no.</label><input data-vf="contact" value="${esc(v.contact)}"></div>
      <div class="field"><label>Status</label><input data-vf="status" value="${esc(v.status)}" list="statusList"></div>
    </div>
    <table class="sheet" style="margin:6px 0"><thead><tr><th>Position</th><th style="width:110px">Required / day</th><th style="width:28px"></th></tr></thead><tbody>
      ${(v.manning || []).map((m, mi) => `<tr data-m="${mi}"><td><select data-mf="position">${posOpts(m.position)}</select></td><td><input data-mf="required" type="number" min="0" max="20" value="${m.required}" style="text-align:center"></td><td><button class="rowbtn" data-vact="delpos" aria-label="Remove position">×</button></td></tr>`).join("")}
    </tbody></table>
    <div class="btnrow"><button class="btn small" data-vact="addpos">+ Position</button>
      <button class="btn small" data-vact="up" ${vi === 0 ? "disabled" : ""}>Move up</button><button class="btn small" data-vact="down" ${vi === S.meta.vessels.length - 1 ? "disabled" : ""}>Move down</button>
      <button class="btn small danger" data-vact="delvessel" style="margin-left:auto">Delete vessel</button></div>
  </div>`).join("") + `<datalist id="statusList"><option>IN SERVICE</option><option>OFF HIRE</option><option>IN DRY DOCK</option><option>STANDBY</option></datalist>`;
}
function bindVessels() {
  $("#addVessel").addEventListener("click", () => {
    S.meta.vessels.push({ id: M.uid("v"), name: "MV NEW VESSEL", client: "", contact: "", status: "IN SERVICE", manning: DEFAULT_POSITIONS.filter((p) => !["MATE / DPO", "MATE - TRAINEE", "OILER / SAILOR - OS"].includes(p)).map((p) => ({ position: p, required: p === "MATE" || p === "SAILOR - OS" ? 2 : 1 })) });
    scheduleSave("meta"); renderVessels(); $("#vesselsList .card:last-child input")?.focus();
  });
  const box = $("#vesselsList");
  box.addEventListener("change", (e) => {
    const card = e.target.closest("[data-vessel]"); if (!card) return;
    const v = S.meta.vessels.find((x) => x.id === card.dataset.vessel);
    if (e.target.dataset.vf) { v[e.target.dataset.vf] = e.target.dataset.vf === "name" ? e.target.value.toUpperCase() : e.target.value; if (e.target.dataset.vf === "name") card.querySelector("h3").textContent = v.name; }
    if (e.target.dataset.mf) { const m = v.manning[+e.target.closest("tr").dataset.m]; m[e.target.dataset.mf] = e.target.dataset.mf === "required" ? Math.max(0, +e.target.value || 0) : e.target.value; }
    scheduleSave("meta"); refreshChecksBadge();
  });
  box.addEventListener("click", async (e) => {
    const b = e.target.closest("[data-vact]"); if (!b) return;
    const card = b.closest("[data-vessel]"); const vi = S.meta.vessels.findIndex((x) => x.id === card.dataset.vessel); const v = S.meta.vessels[vi];
    const a = b.dataset.vact;
    if (a === "addpos") v.manning.push({ position: settings().positions[0] || "", required: 1 });
    if (a === "delpos") v.manning.splice(+b.closest("tr").dataset.m, 1);
    if (a === "up" && vi > 0) [S.meta.vessels[vi - 1], S.meta.vessels[vi]] = [S.meta.vessels[vi], S.meta.vessels[vi - 1]];
    if (a === "down" && vi < S.meta.vessels.length - 1) [S.meta.vessels[vi + 1], S.meta.vessels[vi]] = [S.meta.vessels[vi], S.meta.vessels[vi + 1]];
    if (a === "delvessel") {
      const used = S.week.lines.filter((l) => l.vessel === v.id).length;
      if (!(await dialog("Delete vessel?", `<p>Delete <b>${esc(v.name)}</b>?${used ? ` It has ${used} line(s) on this week's roster – they will move to "Lines with a deleted vessel".` : ""} Past weeks keep their records.</p><p class="hint">To take a vessel out temporarily, set its status to OFF HIRE instead.</p>`, { ok: "Delete", danger: true }))) return;
      S.meta.vessels.splice(vi, 1);
    }
    scheduleSave("meta"); renderVessels(); refreshChecksBadge();
  });
}

// ---------------------------------------------------------------- SETTINGS
function renderSettings() {
  $("#setCompany").value = settings().company || "";
  $("#setPortCaptain").value = settings().portCaptain || "";
  $("#codesTable").innerHTML = `<thead><tr><th style="width:80px">Code</th><th>Meaning</th><th style="width:140px">Label on printout</th><th>Counts toward manning</th><th>Paid</th><th>Note required</th><th style="width:28px"></th></tr></thead><tbody>` +
    settings().codes.map((c, i) => `<tr data-ci="${i}"><td><input data-k="code" value="${esc(c.code)}" style="font-weight:700;text-align:center" ${["W", "ON", "OFF"].includes(c.code) ? "readonly title='Built-in code'" : ""}></td>
      <td><input data-k="meaning" value="${esc(c.meaning)}"></td><td><input data-k="label" value="${esc(c.label)}"></td>
      ${["manning", "paid", "noteRequired"].map((k) => `<td style="text-align:center"><input type="checkbox" data-k="${k}" ${c[k] ? "checked" : ""} style="width:auto;min-height:0"></td>`).join("")}
      <td>${["W", "ON", "OFF"].includes(c.code) ? "" : `<button class="rowbtn" data-cact="del" aria-label="Remove code">×</button>`}</td></tr>`).join("") + "</tbody>";
  $("#positionsTable").innerHTML = `<tbody>` + settings().positions.map((p, i) => `<tr data-pi="${i}"><td><input value="${esc(p)}"></td>
    <td style="width:60px;white-space:nowrap"><button class="rowbtn" data-pact="up" style="width:28px" aria-label="Move up">↑</button><button class="rowbtn" data-pact="down" style="width:28px" aria-label="Move down">↓</button></td>
    <td style="width:28px"><button class="rowbtn" data-pact="del" aria-label="Remove position">×</button></td></tr>`).join("") + "</tbody>";
  $("#storageInfo").innerHTML = S.mode === "firebase"
    ? `Saved to your Firebase project${S.user ? ` as <b>${esc(S.user.email || S.user.name)}</b>` : ""}. Your user ID for the security rules: <code>${esc(S.user?.uid || "")}</code>`
    : `<b>Local test mode</b> – data is saved only in this browser on this computer. Connect Firebase (see README) before relying on it, and download backups regularly.`;
  $("#authRow").hidden = S.mode !== "firebase";
}
function bindSettings() {
  $("#setCompany").addEventListener("input", (e) => { settings().company = e.target.value; scheduleSave("meta"); $("#companyName").textContent = e.target.value; });
  $("#setPortCaptain").addEventListener("input", (e) => { settings().portCaptain = e.target.value; scheduleSave("meta"); });
  $("#codesTable").addEventListener("change", async (e) => {
    const tr = e.target.closest("tr[data-ci]"); if (!tr) return;
    const c = settings().codes[+tr.dataset.ci]; const k = e.target.dataset.k;
    if (k === "code") {
      const v = e.target.value.trim().toUpperCase();
      if (!v || settings().codes.some((x) => x !== c && x.code === v)) { e.target.value = c.code; return; }
      const old = c.code; c.code = v; e.target.value = v;
      S.week.lines.forEach((l) => l.days.forEach((d) => { if (d.code === old) d.code = v; })); scheduleSave("week");
    } else c[k] = e.target.type === "checkbox" ? e.target.checked : e.target.value;
    scheduleSave("meta"); renderHeader(); refreshChecksBadge();
  });
  $("#codesTable").addEventListener("click", (e) => {
    if (!e.target.closest("[data-cact]")) return;
    settings().codes.splice(+e.target.closest("tr").dataset.ci, 1); scheduleSave("meta"); renderSettings(); renderHeader();
  });
  $("#addCode").addEventListener("click", () => { settings().codes.push({ code: "NEW", meaning: "", label: "", manning: false, paid: true, noteRequired: false }); scheduleSave("meta"); renderSettings(); });
  $("#positionsTable").addEventListener("change", (e) => {
    const i = +e.target.closest("tr").dataset.pi; const old = settings().positions[i]; const v = e.target.value.trim().toUpperCase();
    if (!v) { e.target.value = old; return; }
    settings().positions[i] = v;
    S.meta.vessels.forEach((ve) => ve.manning.forEach((m) => { if (m.position === old) m.position = v; }));
    S.week.lines.forEach((l) => { if (l.position === old) l.position = v; });
    S.crew.forEach((c) => { if (c.rank === old) c.rank = v; });
    scheduleSave("meta"); scheduleSave("week"); scheduleSave("crew"); renderSettings();
  });
  $("#positionsTable").addEventListener("click", (e) => {
    const b = e.target.closest("[data-pact]"); if (!b) return;
    const p = settings().positions; const i = +b.closest("tr").dataset.pi;
    if (b.dataset.pact === "up" && i > 0) [p[i - 1], p[i]] = [p[i], p[i - 1]];
    if (b.dataset.pact === "down" && i < p.length - 1) [p[i + 1], p[i]] = [p[i], p[i + 1]];
    if (b.dataset.pact === "del") p.splice(i, 1);
    scheduleSave("meta"); renderSettings();
  });
  $("#addPosition").addEventListener("click", () => { settings().positions.push("NEW POSITION"); scheduleSave("meta"); renderSettings(); });
  $("#backupJson").addEventListener("click", downloadBackup);
  $("#restoreJson").addEventListener("change", async (e) => { const f = e.target.files[0]; e.target.value = ""; if (f) await restoreBackup(await f.text()); });
  $("#signOut").addEventListener("click", () => S.store.signOut());
}

// ---------------------------------------------------------------- backup / restore
async function downloadBackup() {
  await flushSaves();
  const weeks = [];
  for (const id of S.meta.weekIds) { const w = id === S.week.weekEnding ? S.week : await S.store.loadWeek(id); if (w) weeks.push(w); }
  const data = { version: 1, exportedAt: new Date().toISOString(), settings: settings(), vessels: S.meta.vessels, crew: S.crew, weeks };
  download(new Blob([JSON.stringify(data, null, 1)], { type: "application/json" }), `crew-roster-backup-${new Date().toISOString().slice(0, 10)}.json`);
}
export async function importData(data, { silent = false } = {}) {
  if (!data || !Array.isArray(data.crew) || !Array.isArray(data.vessels)) throw new Error("Not a roster backup file");
  if (!silent && !(await dialog("Restore from backup?", `<p>This replaces the crew list, vessels, settings and the weeks in the file (<b>${(data.weeks || []).length}</b> week(s), <b>${data.crew.length}</b> crew, <b>${data.vessels.length}</b> vessels).</p><p class="hint">Download a backup of the current data first if you might need it.</p>`, { ok: "Restore", danger: true }))) return;
  const st = { company: "", portCaptain: "", codes: M.DEFAULT_CODES, positions: DEFAULT_POSITIONS, ...(data.settings || {}) };
  st.codes = st.codes.map((c) => ({ noteRequired: false, label: "", ...c }));
  S.meta = { settings: st, vessels: data.vessels, weekIds: (data.weeks || []).map((w) => w.weekEnding) };
  S.crew = data.crew; reindexCrew();
  for (const w of data.weeks || []) { normaliseWeek(w); await S.store.saveWeek(w); }
  await S.store.saveMeta(S.meta); await S.store.saveCrew(S.crew);
  const latest = [...S.meta.weekIds].sort().pop();
  S.week = null;
  S.tab = "roster";
  if (latest) await openWeek(latest); else await ensureCurrentWeek();
}
async function restoreBackup(text) {
  try { await importData(JSON.parse(text)); }
  catch (e) { await dialog("Couldn't restore", `<p>${esc(e.message)}</p>`, { cancel: "" }); }
}
function download(blob, name) {
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

// ---------------------------------------------------------------- new week / print / export
async function newWeek() {
  const latest = [...S.meta.weekIds].sort().pop();
  if (latest && latest !== S.week.weekEnding) {
    const go = await dialog("Not the latest week", `<p>You're viewing week ending ${M.fmtShort(S.week.weekEnding)}. A new week is always started from the latest week (${M.fmtShort(latest)}).</p>`, { ok: "Go to latest week" });
    if (go) await openWeek(latest); return;
  }
  const { week, carried, removed } = M.startNewWeek(S.week, settings(), S.crewById);
  const notW = carried.filter((c) => c.fill !== "W");
  const ok = await dialog(`Start week ending ${M.fmtShort(week.weekEnding)}?`,
    `<p><b>${carried.length}</b> crew carry over and are pre-filled for all 7 days with their last code – change only the exceptions.</p>
     ${notW.length ? `<p>Carried with a non-working code (note kept): ${notW.map((c) => `<b>${esc(c.name)}</b> (${esc(c.fill)})`).join(", ")}</p>` : ""}
     <p><b>${removed.length}</b> removed${removed.length ? ":" : "."}</p>${removed.length ? `<ul>${removed.map((r) => `<li>${esc(r.name)} – ${esc(r.why)}</li>`).join("")}</ul>` : ""}
     <p class="hint">Week ending ${M.fmtShort(S.week.weekEnding)} stays saved and can be reopened from the week list at the top.</p>`, { ok: "Start new week" });
  if (!ok) return;
  await S.store.saveWeek(S.week);
  S.week = week;
  if (!S.meta.weekIds.includes(week.weekEnding)) S.meta.weekIds.push(week.weekEnding);
  await S.store.saveWeek(S.week); await S.store.saveMeta(S.meta);
  renderAll(); switchTab("roster");
}
async function doPrint(mode) {
  const issues = M.runChecks(S.week, S.meta, S.crewById);
  const blocks = issues.filter((x) => x.sev === "block");
  if (mode === "confirmed" && blocks.length) {
    const ok = await dialog("Not ready to confirm", `<p>There ${blocks.length === 1 ? "is 1 item" : `are ${blocks.length} items`} to fix first:</p><ul>${blocks.slice(0, 12).map((b) => `<li>${esc(b.msg)}</li>`).join("")}${blocks.length > 12 ? "<li>…</li>" : ""}</ul>`, { ok: "Print anyway", cancel: "Go to Checks" });
    if (!ok) { switchTab("checks"); return; }
  }
  const at = new Date().toISOString();
  if (mode === "confirmed") { S.week.status = "confirmed"; S.week.confirmedAt = at; }
  else { if (S.week.status !== "confirmed") S.week.status = "lookahead"; S.week.lookaheadPrintedAt = at; }
  scheduleSave("week"); renderStatus();
  S.printStamp = { mode, at };
  renderBreakdown();
  const title = document.title;
  document.title = `Crew Roster WE ${M.fmtShort(S.week.weekEnding)} - ${mode === "confirmed" ? "CONFIRMED" : "LOOK-AHEAD"}`;
  const restore = () => { document.title = title; S.printStamp = null; renderBreakdown(); };
  window.addEventListener("afterprint", restore, { once: true });
  window.print();
}
async function doExcel() {
  if (!window.ExcelJS) { await dialog("Excel export unavailable", "<p>The Excel library didn't load – check your internet connection and reload the page.</p>", { cancel: "" }); return; }
  const blob = await exportExcel(window.ExcelJS, S.week, S.meta, S.crewById, { stamp: stampFor() });
  download(blob, `Crew Roster WE ${M.fmtShort(S.week.weekEnding)}.xlsx`);
}

// ---------------------------------------------------------------- tabs & boot
function switchTab(t) {
  S.tab = t;
  $$(".tab").forEach((b) => b.classList.toggle("active", b.dataset.tab === t));
  $$(".panel").forEach((p) => p.classList.toggle("active", p.id === "tab-" + t));
  if (t === "roster") renderRoster();
  if (t === "breakdown") renderBreakdown();
  if (t === "checks") renderChecks();
  if (t === "crew") renderCrew();
  if (t === "vessels") renderVessels();
  if (t === "settings") renderSettings();
}
function renderAll() { renderHeader(); switchTab(S.tab); refreshChecksBadge(); }

async function ensureCurrentWeek() {
  const today = new Date(); const iso = new Date(Date.UTC(today.getFullYear(), today.getMonth(), today.getDate())).toISOString().slice(0, 10);
  const we = M.sundayOnOrAfter(iso);
  S.week = M.blankWeek(we, settings());
  if (!S.meta.weekIds.includes(we)) S.meta.weekIds.push(we);
  await S.store.saveWeek(S.week); await S.store.saveMeta(S.meta);
  renderAll();
}

async function boot() {
  S.store = createStore();
  if (S.store.mode === "firebase") {
    $("#signInBtn").addEventListener("click", async () => { try { await S.store.signIn(); } catch (e) { $("#signinMsg").textContent = e.message; } });
  }
  let info;
  try { info = await S.store.init(() => { $("#signin").hidden = false; }); }
  catch (e) { document.body.innerHTML = `<div class="card" style="margin:40px auto;max-width:520px"><h3>Couldn't connect to Firebase</h3><p>${esc(e.message)}</p><p class="hint">Check js/firebase-config.js and your internet connection.</p></div>`; return; }
  $("#signin").hidden = true;
  S.mode = info.mode; S.user = info.user || null;
  if (S.mode === "local") { const b = $("#modeBanner"); b.hidden = false; b.innerHTML = "<b>Local test mode:</b> data is saved only in this browser. Connect your Firebase project (see README) before using it for real, and keep backups."; }

  let loadedMeta;
  try { loadedMeta = await S.store.loadMeta(); }
  catch (e) {
    const denied = /permission|insufficient/i.test(e.message || "");
    document.body.innerHTML = `<div class="card" style="margin:40px auto;max-width:560px"><h3>${denied ? "Access not allowed yet" : "Couldn't load data"}</h3>
      ${denied ? `<p>You're signed in as <b>${esc(S.user?.email || "")}</b>, but the database rules don't allow this account yet.</p>
      <p>Put this user ID into <code>firestore.rules</code> and publish the rules (README step 4):</p><p><code style="font-size:14px">${esc(S.user?.uid || "")}</code></p>` : `<p>${esc(e.message)}</p>`}
      <button class="btn" onclick="location.reload()">Try again</button></div>`;
    return;
  }
  S.meta = loadedMeta || { settings: { company: "", portCaptain: "", codes: M.DEFAULT_CODES, positions: DEFAULT_POSITIONS }, vessels: [], weekIds: [] };
  S.meta.settings.codes = (S.meta.settings.codes || M.DEFAULT_CODES).map((c) => ({ noteRequired: false, label: "", ...c }));
  S.meta.settings.positions ||= DEFAULT_POSITIONS;
  S.meta.weekIds ||= [];
  S.crew = await S.store.loadCrew(); reindexCrew();

  bindRoster(); bindCrew(); bindVessels(); bindSettings();
  $$(".tab").forEach((b) => b.addEventListener("click", () => switchTab(b.dataset.tab)));
  $("#weekSelect").addEventListener("change", (e) => openWeek(e.target.value));
  $("#prevWeek").addEventListener("click", () => { const s = [...S.meta.weekIds].sort(); const i = s.indexOf(S.week.weekEnding); if (i > 0) openWeek(s[i - 1]); });
  $("#nextWeek").addEventListener("click", () => { const s = [...S.meta.weekIds].sort(); const i = s.indexOf(S.week.weekEnding); if (i < s.length - 1) openWeek(s[i + 1]); });
  $("#newWeekBtn").addEventListener("click", newWeek);
  $("#printLookahead").addEventListener("click", () => doPrint("lookahead"));
  $("#printConfirmed").addEventListener("click", () => doPrint("confirmed"));
  $("#exportExcel").addEventListener("click", doExcel);
  $("#showEmptyLines").addEventListener("change", renderBreakdown);
  $("#checksList").addEventListener("click", (e) => { const b = e.target.closest("[data-jump]"); if (b) jumpTo(b.dataset.jump); });
  window.addEventListener("beforeunload", (e) => { if (Object.values(timers).length && $("#saveState").textContent.startsWith("Saving")) { e.preventDefault(); e.returnValue = ""; } });

  const latest = [...S.meta.weekIds].sort().pop();
  if (latest) await openWeek(latest); else await ensureCurrentWeek();
  if (!S.crew.length && !S.meta.vessels.length) switchTab("settings");
  window.__roster = { S, importData, M }; // for testing / console use
}
boot();
