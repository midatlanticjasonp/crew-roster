// Styled Excel export (ExcelJS): "Daily Breakdown" in the familiar layout + a flat "Roster" sheet for payroll maths.
import * as M from "./model.js";

const NAVY = "FF1F3864";
const FILL = { A: "FFF8CBAD", S: "FFFCE4D6", P: "FFFFE699", AL: "FFDDEBF7", L: "FFDDEBF7", CC: "FFBDD7EE", O: "FFE7E6E6", T: "FFE4DFEC", TR: "FFE4DFEC", W: "FFC6EFCE" };
const FONT = { A: "FF9C0006", S: "FFC65911", P: "FF7F6000", AL: "FF1F4E78", L: "FF1F4E78", CC: "FF1F3864", O: "FF404040", T: "FF5B3C88", TR: "FF5B3C88", W: "FF006100" };
const thin = { style: "thin", color: { argb: "FF9AA6B8" } };
const box = { top: thin, left: thin, bottom: thin, right: thin };
const solid = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });

export async function exportExcel(ExcelJS, week, meta, crewById, { stamp } = {}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Crew Roster";
  const dates = M.weekDates(week.weekEnding);
  const cmap = M.codeMap(meta.settings);
  const stampTxt = { draft: "DRAFT", lookahead: "LOOK-AHEAD", confirmed: "CONFIRMED" }[stamp?.mode || "draft"];

  // ---- Daily Breakdown ----
  const ws = wb.addWorksheet("Daily Breakdown", {
    pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.35, right: 0.35, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 } },
    views: [{ showGridLines: false }],
  });
  ws.columns = [{ width: 22 }, { width: 22 }, ...dates.map(() => ({ width: 19 }))];
  let r = 1;
  const bd = M.buildBreakdown(week, meta, crewById, { showEmpty: true });
  const sig = stamp?.mode === "confirmed" && meta.settings.signature?.dataUrl;
  let sigId = null, sigW = 0, sigH = 0;
  if (sig) {
    sigId = wb.addImage({ base64: sig, extension: "png" });
    const dim = await new Promise((res) => { const i = new Image(); i.onload = () => res([i.naturalWidth, i.naturalHeight]); i.onerror = () => res([300, 100]); i.src = sig; });
    sigH = 30; sigW = Math.round((dim[0] / dim[1]) * sigH);
  }
  bd.forEach(({ vessel, groups, notes }, vi) => {
    if (vi > 0) { ws.getRow(r - 1).addPageBreak(); }
    ws.mergeCells(r, 1, r, 9);
    Object.assign(ws.getCell(r, 1), { value: stamp?.mode === "confirmed" ? "VESSELS - WEEKLY ROSTER" : `VESSELS - WEEKLY ROSTER   [${stampTxt}]` });
    ws.getCell(r, 1).font = { bold: true, size: 14, color: { argb: NAVY } }; ws.getCell(r, 1).alignment = { horizontal: "center" };
    r += 2;
    ws.getCell(r, 1).value = `COMPANY: ${meta.settings.company || ""}`;
    ws.getCell(r, 4).value = "Week Ending:"; ws.getCell(r, 5).value = M.fmtShort(week.weekEnding);
    ws.getCell(r, 7).value = "Port Captain:"; ws.getCell(r, 8).value = week.portCaptain || "";
    if (sigId !== null) ws.addImage(sigId, { tl: { col: 8.05, row: r - 1.3 }, ext: { width: sigW, height: sigH }, editAs: "oneCell" });
    [1, 4, 5, 7, 8].forEach((c) => (ws.getCell(r, c).font = { bold: true }));
    r += 2;
    const hdr = ["VESSEL", "CREW", ...dates.map((d, i) => M.dayHeader(d, i))];
    hdr.forEach((h, i) => { const c = ws.getCell(r, i + 1); c.value = h; c.font = { bold: true, color: { argb: "FFFFFFFF" } }; c.fill = solid(NAVY); c.alignment = { horizontal: "center", vertical: "middle" }; c.border = box; });
    ws.getRow(r).height = 24;
    r++;
    const top = r;
    for (const g of groups) {
      for (const [ri, row] of g.rows.entries()) {
        const pc = ws.getCell(r, 2);
        pc.value = g.rows.length < 2 ? g.position : /\d$/.test(g.position) ? `${g.position} (${ri + 1})` : `${g.position} ${ri + 1}`;
        pc.font = { bold: true, size: 10 }; pc.alignment = { vertical: "middle", horizontal: "left" };
        row.forEach((ents, di) => {
          const c = ws.getCell(r, 3 + di);
          c.value = ents.map((e) => e.name + (e.text ? "\n" + e.text : "")).join("\n");
          c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
          c.font = { bold: true, size: 9 }; c.border = box;
          const key = ents.length === 1 ? ents[0].key : ents.some((e) => e.key === "CC") ? "CC" : "";
          if (key && key !== "W" && FILL[key]) { c.fill = solid(FILL[key]); c.font = { bold: true, size: 9, color: { argb: FONT[key] } }; }
        });
        const lines = Math.max(1, ...row.map((ents) => ents.reduce((n, e) => n + 1 + (e.text ? 1 : 0), 0)));
        ws.getRow(r).height = Math.max(24, lines * 12 + 6);
        ws.getCell(r, 2).border = box;
        r++;
      }
    }
    if (r === top) { for (let c = 2; c <= 9; c++) ws.getCell(r, c).border = box; r++; }
    if (r - 1 > top) ws.mergeCells(top, 1, r - 1, 1);
    const vc = ws.getCell(top, 1);
    vc.value = [vessel.name.replace(/^MV\s+/i, ""), "", vessel.client || vessel.status ? `${vessel.status || "IN SERVICE"}${vessel.client ? " WITH " + vessel.client : ""}` : "", "", vessel.contact || ""].join("\n").trim();
    vc.font = { bold: true, size: 11 }; vc.alignment = { horizontal: "center", vertical: "top", wrapText: true }; vc.border = box;
    r++;
    ws.getCell(r, 1).value = "Notes:"; ws.getCell(r, 1).font = { bold: true };
    const noteLines = [...(week.vesselNotes?.[vessel.id] || "").split("\n").map((s) => s.trim()).filter(Boolean), ...notes.map((n) => `${n.name} – ${n.note}`)];
    if (!noteLines.length) r++;
    for (const t of noteLines) { ws.mergeCells(r, 2, r, 9); const c = ws.getCell(r, 2); c.value = t; c.alignment = { wrapText: true, vertical: "top" }; c.font = { size: 9 }; ws.getRow(r).height = Math.max(15, Math.ceil(t.length / 150) * 13); r++; }
    r++;
    ws.getCell(r, 1).value = `Prepared by: ${week.preparedBy || "____________"}`;
    ws.getCell(r, 4).value = `Checked by: ${week.checkedBy || "____________"}`;
    ws.getCell(r, 7).value = "Port Captain: ____________";
    if (sigId !== null) ws.addImage(sigId, { tl: { col: 7.25, row: r - 1.5 }, ext: { width: sigW, height: sigH }, editAs: "oneCell" });
    for (const c of [1, 4, 7]) ws.getCell(r, c).font = { size: 9, color: { argb: "FF595959" } };
    r += 3;
  });

  // ---- Roster (flat, one line per person) ----
  const rs = wb.addWorksheet("Roster", { views: [{ state: "frozen", xSplit: 3, ySplit: 1 }] });
  const head = ["Vessel", "Position", "Crew", "Full name", "Crew ID", ...dates.map((d, i) => `${M.DAYS[i]} ${M.fmtShort(d).slice(0, 5)}`), "Days on board", "Paid days", "Relieving", "Note"];
  rs.addRow(head);
  rs.getRow(1).eachCell((c) => { c.font = { bold: true, color: { argb: "FFFFFFFF" } }; c.fill = solid(NAVY); c.alignment = { horizontal: "center", vertical: "middle", wrapText: true }; c.border = box; });
  rs.getRow(1).height = 30;
  const vName = new Map(meta.vessels.map((v) => [v.id, v.name]));
  for (const v of [...meta.vessels, { id: null }]) {
    const lines = v.id ? M.sortLines(week.lines.filter((l) => l.vessel === v.id), v, meta.settings) : week.lines.filter((l) => !vName.has(l.vessel));
    for (const l of lines) {
      if (!l.crewId && !M.hasDays(l)) continue;
      const c = crewById.get(l.crewId) || {};
      const st = M.lineStats(l, cmap);
      const row = rs.addRow([vName.get(l.vessel) || "?", l.position, c.name || "", c.fullName || "", c.id || "",
        ...l.days.map((d) => (d.code ? d.code + (d.time ? " " + d.time : "") : "")), st.onBoard, st.paid, crewById.get(l.relievingCrewId)?.name || "", l.note || ""]);
      row.eachCell({ includeEmpty: true }, (cell, ci) => {
        cell.border = box;
        if (ci >= 6 && ci <= 12) {
          cell.alignment = { horizontal: "center" };
          const code = l.days[ci - 6].code; const key = M.isCC(code) ? "CC" : M.cssKey(code);
          if (code && FILL[key]) { cell.fill = solid(FILL[key]); cell.font = { bold: true, color: { argb: FONT[key] } }; }
        }
        if (ci === 13 || ci === 14) cell.alignment = { horizontal: "center" };
      });
    }
  }
  rs.columns = [{ width: 22 }, { width: 20 }, { width: 18 }, { width: 22 }, { width: 10 }, ...dates.map(() => ({ width: 11 })), { width: 9 }, { width: 8 }, { width: 18 }, { width: 60 }];
  rs.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: head.length } };
  rs.pageSetup = { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

  // ---- Codes key ----
  const ks = wb.addWorksheet("Codes");
  ks.addRow(["Code", "Meaning", "Counts toward manning", "Paid"]).eachCell((c) => { c.font = { bold: true, color: { argb: "FFFFFFFF" } }; c.fill = solid(NAVY); });
  (meta.settings.codes || []).forEach((c) => ks.addRow([c.code, c.meaning, c.manning ? "Y" : "N", c.paid ? "Y" : "N"]));
  ks.columns = [{ width: 8 }, { width: 50 }, { width: 22 }, { width: 8 }];

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

// ---- Crew list (Crew tab) ----
export async function exportCrewExcel(ExcelJS, crew, meta, week, { onWeek } = {}) {
  const wb = new ExcelJS.Workbook(); wb.creator = "Crew Roster";
  const ws = wb.addWorksheet("Crew", { views: [{ state: "frozen", xSplit: 2, ySplit: 1 }] });
  const head = ["Crew ID", "Roster name", "Full name", "Rank", "Phone", "Email", "CoC / Licence expiry", "Medical expiry", "Offshore safety expiry", "Certificates", `On roster WE ${M.fmtShort(week.weekEnding)}`, "Active", "Notes"];
  ws.addRow(head);
  ws.getRow(1).eachCell((c) => { c.font = { bold: true, color: { argb: "FFFFFFFF" } }; c.fill = solid(NAVY); c.alignment = { horizontal: "center", vertical: "middle", wrapText: true }; c.border = box; });
  ws.getRow(1).height = 32;
  const toDate = (iso) => { if (!iso) return null; const [y, m, d] = iso.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)); };
  const certFill = { ok: ["FFC6EFCE", "FF006100"], due: ["FFFFE699", "FF7F6000"], exp: ["FFF8CBAD", "FF9C0006"] };
  const sorted = [...crew].sort((a, b) => (a.active === false) - (b.active === false) || String(a.name).localeCompare(String(b.name)));
  for (const c of sorted) {
    const cs = M.certStatus(c, week.weekEnding);
    const row = ws.addRow([c.id || "", c.name || "", c.fullName || "", c.rank || "", c.phone || "", c.email || "",
      toDate(c.cocExpiry), toDate(c.medicalExpiry), toDate(c.offshoreExpiry), cs.text, (onWeek?.get(c.id) || "").replace(/^MV /, ""),
      c.active === false ? "No" : "Yes", c.notes || ""]);
    row.eachCell({ includeEmpty: true }, (cell, ci) => {
      cell.border = box; cell.alignment = { vertical: "middle", wrapText: ci === 13 };
      if (ci >= 7 && ci <= 9) { cell.numFmt = "dd/mm/yyyy"; cell.alignment = { horizontal: "center", vertical: "middle" }; }
      if (ci === 10 && certFill[cs.cls]) { cell.fill = solid(certFill[cs.cls][0]); cell.font = { bold: true, color: { argb: certFill[cs.cls][1] } }; cell.alignment = { horizontal: "center", vertical: "middle" }; }
      if (ci === 12) cell.alignment = { horizontal: "center", vertical: "middle" };
    });
    if (c.active === false) row.font = { color: { argb: "FF7F7F7F" } };
    row.getCell(2).font = { bold: true, color: { argb: c.active === false ? "FF7F7F7F" : "FF000000" } };
  }
  ws.columns = [{ width: 10 }, { width: 18 }, { width: 24 }, { width: 20 }, { width: 14 }, { width: 30 }, { width: 13 }, { width: 13 }, { width: 13 }, { width: 22 }, { width: 18 }, { width: 8 }, { width: 50 }];
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: head.length } };
  ws.pageSetup = { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: "1:1" };
  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}
