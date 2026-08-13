import JSZip from "jszip";
import { toExcelSerial, formatMDY } from "./payPeriods.js";
import { normalizeEmployeeId } from "./employeeId.js";

// ---------------------------------------------------------------------------
// How this works, in plain terms
// ---------------------------------------------------------------------------
// An .xlsx file is just a zip full of XML files. Rather than fully parsing
// the workbook into a generic object model (which is how the `xlsx` library
// works, and which loses the template's styling/dropdowns/merges on the way
// out), this module edits the *actual XML* directly - the same technique
// the reference OSY/ISY app used. That means every tab this app produces
// keeps the exact fonts, colors, borders, merged cells, and dropdown lists
// from Contractor_Timesheet_PY26.xlsx, because it IS that file's XML,
// lightly edited.
//
// Every employee gets ONE sheet/tab, and every pay period they have hours
// for is a block stacked on that same tab. The first pay period on a tab
// reuses the template's own built-in day-grid (rows 10-23) - including its
// original formulas, so Excel keeps auto-calculating totals exactly like
// the blank template did. Every pay period after that is a freshly-built
// block, appended below the last one, styled to match but with plain
// computed values instead of formulas (simpler and just as correct, since
// we already know the totals).
//
// Where an appended block starts is always computed dynamically - one row
// past whatever the highest row number already on the tab is - rather than
// a fixed offset. That matters because the template itself already has
// content below the first block (an EMPLOYEE SIGNATURE / MANAGER SIGNATURE
// line at rows 28-29 and 34-35): a fixed "next block starts at row 25"
// offset would land a second block's day-rows directly on top of those
// signature rows, producing two <row> elements with the same row number in
// the sheet XML - which is what was silently corrupting/"cutting off" hours
// on multi-period tabs before this fix.
//
// A small marker is stored in cells A1/A2 of every tab (invisible in normal
// use - column A isn't part of the printed/visible template) recording the
// employee ID and which pay periods have already been written to that tab.
// That's what makes it safe to run this tool again next pay period without
// duplicating anything: on the next run, every tab in the "ongoing"
// workbook is scanned for its marker, and matching is done by Employee ID
// with leading zeros stripped (some report exports zero-pad IDs and some
// don't - see lib/employeeId.js).
// ---------------------------------------------------------------------------

const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";

// Shared-string table indices baked into Contractor_Timesheet_PY26.xlsx.
// (See lib/fundMap.js for why these particular labels exist.)
const SHARED = {
  DAY: 4,
  WEDNESDAY: 5,
  THURSDAY: 6,
  FRIDAY: 7,
  MONDAY: 8,
  TUESDAY: 9,
  HOLIDAY: 10,
  LEAVE_TIME: 11,
  TOTAL_HOURS: 12,
  TOTAL_HOURS_COLON: 13,
  EARN: 17,
  "WIOA A": 18,
  "WIOA DW": 19,
  "WIOA Y": 20,
  BST: 21,
  VEAP: 22,
  PTO: 23,
  "TANF Y": 24,
  SNAP: 30,
  OSO: 31,
  ComOut: 32,
  HOURS_WORKED_BCWDB: 33,
  HOURS_WORKED_NONBCWDB: 34,
  "WIOA PWE": 35,
  "UWBC RSD": 36,
  "UWBC GM": 37,
  RSD: 38,
  "Business Services": 39,
};

const MARKER_PREFIX = "MERGER_META:";

// The BCWDB funding-stream dropdown's source range on Sheet1 (Q13:Q24,
// referenced by both the template's own G11:J11 validation and every
// appended block's G:J validation). Extended from $Q$13:$Q$23 to
// $Q$13:$Q$24 when "Business Services" was added at Q24.
const BCWDB_DROPDOWN_RANGE = "$Q$13:$Q$24";
const NON_BCWDB_DROPDOWN_RANGE = "$U$13:$U$17";

const DAY_NAME_CYCLE = [
  SHARED.MONDAY,
  SHARED.TUESDAY,
  SHARED.WEDNESDAY,
  SHARED.THURSDAY,
  SHARED.FRIDAY,
  SHARED.MONDAY,
  SHARED.TUESDAY,
  SHARED.WEDNESDAY,
  SHARED.THURSDAY,
  SHARED.FRIDAY,
];

// Per-column style ids for a cloned (non-formula) block, lifted directly
// from the template's own rows 12 (first day), 13-20 (middle days), and 21
// (last day), plus row 22 (totals). See notes in the README for how these
// were captured.
const DAY_ROW_STYLE = {
  first: { B: 18, C: 19, D: 13, E: 23, F: 24, G: 23, H: 24, I: 24, J: 25, K: 23, L: 24, M: 25 },
  middle: { B: 1, C: 6, D: 14, E: 26, F: 27, G: 26, H: 27, I: 27, J: 28, K: 26, L: 27, M: 28 },
  last: { B: 2, C: 7, D: 16, E: 32, F: 33, G: 32, H: 33, I: 33, J: 34, K: 32, L: 33, M: 34 },
};
const HEADER_ROW_STYLE = { E: 58, F: 59, G: 58, H: 60, I: 60, J: 59, K: 58, L: 60, M: 59 };
const SUBHEADER_ROW_STYLE = { B: 54, C: 55, D: 12, E: 3, F: 4, G: 35, H: 36, I: 36, J: 37, K: 35, L: 36, M: 37 };
const TOTAL_ROW_STYLE = { B: 53, C: 53, D: 10, E: 10, F: 10, G: 10, H: 10, I: 10, J: 10, K: 10, L: 10, M: 10 };
const LABEL_ROW_STYLE = 5; // reused from "PAY PERIOD START DATE:" - bold section label

// ---------------------------------------------------------------------------
// Small XML helpers
// ---------------------------------------------------------------------------

function parseXml(text) {
  return new DOMParser().parseFromString(text, "application/xml");
}
function serializeXml(doc) {
  return new XMLSerializer().serializeToString(doc);
}
function el(doc, name) {
  return doc.createElementNS(NS_MAIN, name);
}
function findRow(sheetData, rowNum) {
  for (const row of Array.from(sheetData.getElementsByTagNameNS(NS_MAIN, "row"))) {
    if (row.getAttribute("r") === String(rowNum)) return row;
  }
  return null;
}
function findCell(row, colLetter) {
  if (!row) return null;
  for (const c of Array.from(row.getElementsByTagNameNS(NS_MAIN, "c"))) {
    if (c.getAttribute("r") === `${colLetter}${row.getAttribute("r")}`) return c;
  }
  return null;
}
/** Highest row number currently present anywhere in the sheet (including rows the template ships with, like the signature lines below row 23). Used to find a safe, collision-free place to append a new block, rather than a fixed offset. */
function maxRowNumber(sheetData) {
  let max = 0;
  for (const row of Array.from(sheetData.getElementsByTagNameNS(NS_MAIN, "row"))) {
    const n = parseInt(row.getAttribute("r"), 10);
    if (!Number.isNaN(n) && n > max) max = n;
  }
  return max;
}
function setCell(doc, rowEl, colLetter, rowNum, { style, sharedIndex, inlineText, number } = {}) {
  const ref = `${colLetter}${rowNum}`;
  const c = el(doc, "c");
  c.setAttribute("r", ref);
  if (style != null) c.setAttribute("s", String(style));
  if (sharedIndex != null) {
    c.setAttribute("t", "s");
    const v = el(doc, "v");
    v.textContent = String(sharedIndex);
    c.appendChild(v);
  } else if (inlineText != null) {
    c.setAttribute("t", "inlineStr");
    const is = el(doc, "is");
    const t = el(doc, "t");
    t.textContent = inlineText;
    is.appendChild(t);
    c.appendChild(is);
  } else if (number != null) {
    const v = el(doc, "v");
    v.textContent = String(number);
    c.appendChild(v);
  }
  rowEl.appendChild(c);
  return c;
}
function setExistingCellValue(doc, sheetData, colLetter, rowNum, { number, inlineText, sharedIndex } = {}) {
  const row = findRow(sheetData, rowNum);
  if (!row) return;
  let c = findCell(row, colLetter);
  if (!c) {
    c = el(doc, "c");
    c.setAttribute("r", `${colLetter}${rowNum}`);
    row.appendChild(c);
  }
  // Clear any existing formula/value children (but keep the style attribute).
  Array.from(c.childNodes).forEach((n) => c.removeChild(n));
  c.removeAttribute("t");
  if (sharedIndex != null) {
    c.setAttribute("t", "s");
    const v = el(doc, "v");
    v.textContent = String(sharedIndex);
    c.appendChild(v);
  } else if (inlineText != null) {
    c.setAttribute("t", "inlineStr");
    const is = el(doc, "is");
    const t = el(doc, "t");
    t.textContent = inlineText;
    is.appendChild(t);
    c.appendChild(is);
  } else if (number != null) {
    const v = el(doc, "v");
    v.textContent = String(number);
    c.appendChild(v);
  }
}

function ensureRowMarkerCells(doc, sheetData) {
  // Rows 1 and 2 always exist in the template (title/contractor rows).
  // We add cells in column A (unused by the visible template) to store
  // machine-readable metadata.
  for (const rowNum of [1, 2]) {
    if (!findRow(sheetData, rowNum)) {
      const row = el(doc, "row");
      row.setAttribute("r", String(rowNum));
      sheetData.appendChild(row);
    }
  }
}

function writeMarker(doc, sheetData, employeeId, meta) {
  ensureRowMarkerCells(doc, sheetData);
  setExistingCellValue(doc, sheetData, "A", 1, { inlineText: `EMP_ID:${employeeId}` });
  setExistingCellValue(doc, sheetData, "A", 2, {
    inlineText: `${MARKER_PREFIX}${JSON.stringify(meta)}`,
  });
}

function readMarker(sheetData) {
  const row2 = findRow(sheetData, 2);
  const cellA2 = findCell(row2, "A");
  if (!cellA2) return null;
  const t = cellA2.getElementsByTagNameNS(NS_MAIN, "t")[0];
  const text = t ? t.textContent : "";
  if (!text.startsWith(MARKER_PREFIX)) return null;
  try {
    return JSON.parse(text.slice(MARKER_PREFIX.length));
  } catch {
    return null;
  }
}

function readEmployeeIdMarker(sheetData) {
  const row1 = findRow(sheetData, 1);
  const cellA1 = findCell(row1, "A");
  if (!cellA1) return null;
  const t = cellA1.getElementsByTagNameNS(NS_MAIN, "t")[0];
  const text = t ? t.textContent : "";
  return text.startsWith("EMP_ID:") ? normalizeEmployeeId(text.slice("EMP_ID:".length)) : null;
}

// ---------------------------------------------------------------------------
// Filling the template's own block (rows 10-23) for an employee's first
// pay period on a tab. We deliberately leave all of the block's original
// formulas untouched - only the plain input cells get written. The marker
// (which employee, which periods) is written once by the caller after all
// of this run's periods for the tab are done, not per-block.
// ---------------------------------------------------------------------------

function fillBlock1(doc, sheetData, { employeeName, period, assignment, managerName }) {
  // Header fields.
  setExistingCellValue(doc, sheetData, "C", 7, { inlineText: employeeName }); // EMPLOYEE:
  setExistingCellValue(doc, sheetData, "C", 4, { number: toExcelSerial(period.start) }); // PAY PERIOD START DATE
  if (managerName) {
    setExistingCellValue(doc, sheetData, "C", 8, { inlineText: managerName }); // MANAGER:
  }

  writeFundHeaderRow(doc, sheetData, 11, assignment);
  writeDayRows(doc, sheetData, 12, period, assignment);
  refreshFormulaCachedValues(doc, sheetData, period, assignment);
}

// Block 1 keeps its original template formulas (dates chained off C4, SUM
// totals, percentages) rather than us re-deriving them - but the formulas'
// *cached* <v> values are still whatever the blank template had (mostly 0,
// or #DIV/0! for the percentage row). Most spreadsheet apps recalculate
// formulas automatically on open (and we also set fullCalcOnLoad="1" on the
// workbook to ask for that), but not every tool that might read this file
// does - so we also write the correct value in directly, alongside the
// formula, exactly like Excel itself does when it saves a file. That way
// the numbers are right even before/without a recalculation pass.
function refreshFormulaCachedValues(doc, sheetData, period, assignment) {
  const bcwdbCols = ["G", "H", "I", "J"];
  const nonBcwdbCols = ["K", "L", "M"];
  const dayTotal = (i) =>
    (period.holidayByDay[i] || 0) +
    (period.ptoByDay[i] || 0) +
    assignment.bcwdb.reduce((s, f) => s + (f.byDay[i] || 0), 0) +
    assignment.nonBcwdb.reduce((s, f) => s + (f.byDay[i] || 0), 0);
  const sumBy = (arr) => (arr || []).reduce((s, n) => s + (n || 0), 0);

  for (let i = 0; i < 10; i++) {
    const rowNum = 12 + i;
    setFormulaCachedValue(sheetData, "C", rowNum, toExcelSerial(period.days[i]));
    setFormulaCachedValue(sheetData, "D", rowNum, dayTotal(i));
  }

  const colTotal = (col) => {
    if (col === "D")
      return (
        sumBy(period.holidayByDay) +
        sumBy(period.ptoByDay) +
        assignment.bcwdb.reduce((s, f) => s + sumBy(f.byDay), 0) +
        assignment.nonBcwdb.reduce((s, f) => s + sumBy(f.byDay), 0)
      );
    if (col === "E") return sumBy(period.holidayByDay);
    if (col === "F") return sumBy(period.ptoByDay);
    const bi = bcwdbCols.indexOf(col);
    if (bi >= 0) return assignment.bcwdb[bi] ? sumBy(assignment.bcwdb[bi].byDay) : 0;
    const ni = nonBcwdbCols.indexOf(col);
    if (ni >= 0) return assignment.nonBcwdb[ni] ? sumBy(assignment.nonBcwdb[ni].byDay) : 0;
    return 0;
  };
  ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"].forEach((col) => {
    setFormulaCachedValue(sheetData, col, 22, colTotal(col));
  });

  // Row 23 percentages: hours in each fund column / total non-leave hours worked.
  const workedTotal = colTotal("D") - colTotal("E") - colTotal("F");
  ["G", "H", "I", "J", "K", "L", "M"].forEach((col) => {
    const pct = workedTotal > 0 ? colTotal(col) / workedTotal : 0;
    setFormulaCachedValue(sheetData, col, 23, pct, { clearErrorType: true });
  });
}

function setFormulaCachedValue(sheetData, colLetter, rowNum, value, { clearErrorType } = {}) {
  const row = findRow(sheetData, rowNum);
  const c = findCell(row, colLetter);
  if (!c) return; // leave any cell we don't recognize alone
  if (clearErrorType) c.removeAttribute("t");
  let v = c.getElementsByTagNameNS(NS_MAIN, "v")[0];
  if (!v) {
    v = c.ownerDocument.createElementNS(NS_MAIN, "v");
    c.appendChild(v);
  }
  v.textContent = String(value);
}

function writeFundHeaderRow(doc, sheetData, rowNum, assignment) {
  const bcwdbCols = ["G", "H", "I", "J"];
  const nonBcwdbCols = ["K", "L", "M"];
  assignment.bcwdb.forEach((fund, i) => {
    if (i >= bcwdbCols.length) return;
    writeFundLabelCell(doc, sheetData, bcwdbCols[i], rowNum, fund.label);
  });
  assignment.nonBcwdb.forEach((fund, i) => {
    if (i >= nonBcwdbCols.length) return;
    writeFundLabelCell(doc, sheetData, nonBcwdbCols[i], rowNum, fund.label);
  });
}

function writeFundLabelCell(doc, sheetData, colLetter, rowNum, label) {
  const sharedIndex = SHARED[label];
  if (sharedIndex != null) {
    setExistingCellValue(doc, sheetData, colLetter, rowNum, { sharedIndex });
  } else {
    setExistingCellValue(doc, sheetData, colLetter, rowNum, { inlineText: label });
  }
}

function writeDayRows(doc, sheetData, startRow, period, assignment) {
  const bcwdbCols = ["G", "H", "I", "J"];
  const nonBcwdbCols = ["K", "L", "M"];

  for (let i = 0; i < 10; i++) {
    const rowNum = startRow + i;

    if (period.holidayByDay[i]) {
      setExistingCellValue(doc, sheetData, "E", rowNum, { number: period.holidayByDay[i] });
    }
    if (period.ptoByDay[i]) {
      setExistingCellValue(doc, sheetData, "F", rowNum, { number: period.ptoByDay[i] });
    }
    assignment.bcwdb.forEach((fund, idx) => {
      if (idx >= bcwdbCols.length) return;
      if (fund.byDay[i]) {
        setExistingCellValue(doc, sheetData, bcwdbCols[idx], rowNum, { number: fund.byDay[i] });
      }
    });
    assignment.nonBcwdb.forEach((fund, idx) => {
      if (idx >= nonBcwdbCols.length) return;
      if (fund.byDay[i]) {
        setExistingCellValue(doc, sheetData, nonBcwdbCols[idx], rowNum, { number: fund.byDay[i] });
      }
    });
  }
}

// ---------------------------------------------------------------------------
// Appending a brand-new (non-formula) block below the last one on a tab, for
// an employee's 2nd+ pay period. `labelRow` is computed by the caller as
// one row past the current highest row number on the sheet, so this never
// collides with either the template's own trailing content (the signature
// lines at rows 28-29/34-35) or an earlier appended block.
// ---------------------------------------------------------------------------

function appendBlock(doc, sheetData, { period, assignment, labelRow }) {
  const headerRow = labelRow + 1;
  const subheaderRow = labelRow + 2;
  const daysStart = labelRow + 3;
  const totalRow = labelRow + 13;

  // --- label row ---
  const rowLabel = el(doc, "row");
  rowLabel.setAttribute("r", String(labelRow));
  sheetData.appendChild(rowLabel);
  setCell(doc, rowLabel, "B", labelRow, {
    style: LABEL_ROW_STYLE,
    inlineText: `Pay Period: ${formatMDY(period.start)} - ${formatMDY(period.end)}`,
  });

  // --- header row (LEAVE TIME / HOURS WORKED BCWDB / HOURS WORKED NON-BCWDB) ---
  const rowHeader = el(doc, "row");
  rowHeader.setAttribute("r", String(headerRow));
  sheetData.appendChild(rowHeader);
  setCell(doc, rowHeader, "E", headerRow, { style: HEADER_ROW_STYLE.E, sharedIndex: SHARED.LEAVE_TIME });
  setCell(doc, rowHeader, "F", headerRow, { style: HEADER_ROW_STYLE.F });
  setCell(doc, rowHeader, "G", headerRow, { style: HEADER_ROW_STYLE.G, sharedIndex: SHARED.HOURS_WORKED_BCWDB });
  setCell(doc, rowHeader, "H", headerRow, { style: HEADER_ROW_STYLE.H });
  setCell(doc, rowHeader, "I", headerRow, { style: HEADER_ROW_STYLE.I });
  setCell(doc, rowHeader, "J", headerRow, { style: HEADER_ROW_STYLE.J });
  setCell(doc, rowHeader, "K", headerRow, { style: HEADER_ROW_STYLE.K, sharedIndex: SHARED.HOURS_WORKED_NONBCWDB });
  setCell(doc, rowHeader, "L", headerRow, { style: HEADER_ROW_STYLE.L });
  setCell(doc, rowHeader, "M", headerRow, { style: HEADER_ROW_STYLE.M });

  // --- subheader row (DAY / TOTAL HOURS / HOLIDAY / PTO / fund labels) ---
  const rowSub = el(doc, "row");
  rowSub.setAttribute("r", String(subheaderRow));
  sheetData.appendChild(rowSub);
  setCell(doc, rowSub, "B", subheaderRow, { style: SUBHEADER_ROW_STYLE.B, sharedIndex: SHARED.DAY });
  setCell(doc, rowSub, "C", subheaderRow, { style: SUBHEADER_ROW_STYLE.C });
  setCell(doc, rowSub, "D", subheaderRow, { style: SUBHEADER_ROW_STYLE.D, sharedIndex: SHARED.TOTAL_HOURS });
  setCell(doc, rowSub, "E", subheaderRow, { style: SUBHEADER_ROW_STYLE.E, sharedIndex: SHARED.HOLIDAY });
  setCell(doc, rowSub, "F", subheaderRow, { style: SUBHEADER_ROW_STYLE.F, sharedIndex: SHARED.PTO });
  const bcwdbCols = ["G", "H", "I", "J"];
  const nonBcwdbCols = ["K", "L", "M"];
  bcwdbCols.forEach((col, i) => {
    const fund = assignment.bcwdb[i];
    const c = setCell(doc, rowSub, col, subheaderRow, { style: SUBHEADER_ROW_STYLE[col] });
    if (fund) writeFundLabelIntoCell(doc, c, fund.label);
  });
  nonBcwdbCols.forEach((col, i) => {
    const fund = assignment.nonBcwdb[i];
    const c = setCell(doc, rowSub, col, subheaderRow, { style: SUBHEADER_ROW_STYLE[col] });
    if (fund) writeFundLabelIntoCell(doc, c, fund.label);
  });

  // --- 10 day rows ---
  for (let i = 0; i < 10; i++) {
    const rowNum = daysStart + i;
    const styleSet = i === 0 ? DAY_ROW_STYLE.first : i === 9 ? DAY_ROW_STYLE.last : DAY_ROW_STYLE.middle;
    const row = el(doc, "row");
    row.setAttribute("r", String(rowNum));
    sheetData.appendChild(row);

    setCell(doc, row, "B", rowNum, { style: styleSet.B, sharedIndex: DAY_NAME_CYCLE[i] });
    setCell(doc, row, "C", rowNum, { style: styleSet.C, number: toExcelSerial(period.days[i]) });

    const dayTotal =
      (period.holidayByDay[i] || 0) +
      (period.ptoByDay[i] || 0) +
      assignment.bcwdb.reduce((s, f) => s + (f.byDay[i] || 0), 0) +
      assignment.nonBcwdb.reduce((s, f) => s + (f.byDay[i] || 0), 0);
    setCell(doc, row, "D", rowNum, { style: styleSet.D, number: dayTotal });
    setCell(doc, row, "E", rowNum, { style: styleSet.E, number: period.holidayByDay[i] || undefined });
    setCell(doc, row, "F", rowNum, { style: styleSet.F, number: period.ptoByDay[i] || undefined });
    bcwdbCols.forEach((col, idx) => {
      const fund = assignment.bcwdb[idx];
      setCell(doc, row, col, rowNum, { style: styleSet[col], number: fund ? fund.byDay[i] || undefined : undefined });
    });
    nonBcwdbCols.forEach((col, idx) => {
      const fund = assignment.nonBcwdb[idx];
      setCell(doc, row, col, rowNum, { style: styleSet[col], number: fund ? fund.byDay[i] || undefined : undefined });
    });
  }

  // --- total row ---
  const rowTotal = el(doc, "row");
  rowTotal.setAttribute("r", String(totalRow));
  sheetData.appendChild(rowTotal);
  setCell(doc, rowTotal, "B", totalRow, { style: TOTAL_ROW_STYLE.B, sharedIndex: SHARED.TOTAL_HOURS_COLON });
  setCell(doc, rowTotal, "C", totalRow, { style: TOTAL_ROW_STYLE.C });
  const colTotal = (col) => {
    if (col === "D")
      return (
        sumBy(period.holidayByDay) +
        sumBy(period.ptoByDay) +
        assignment.bcwdb.reduce((s, f) => s + sumBy(f.byDay), 0) +
        assignment.nonBcwdb.reduce((s, f) => s + sumBy(f.byDay), 0)
      );
    if (col === "E") return sumBy(period.holidayByDay);
    if (col === "F") return sumBy(period.ptoByDay);
    const bi = bcwdbCols.indexOf(col);
    if (bi >= 0) return assignment.bcwdb[bi] ? sumBy(assignment.bcwdb[bi].byDay) : 0;
    const ni = nonBcwdbCols.indexOf(col);
    if (ni >= 0) return assignment.nonBcwdb[ni] ? sumBy(assignment.nonBcwdb[ni].byDay) : 0;
    return 0;
  };
  ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"].forEach((col) => {
    const total = colTotal(col);
    setCell(doc, rowTotal, col, totalRow, { style: TOTAL_ROW_STYLE[col], number: total || undefined });
  });

  addMergeCell(doc, `B${labelRow}:M${labelRow}`);
  addMergeCell(doc, `E${headerRow}:F${headerRow}`);
  addMergeCell(doc, `G${headerRow}:J${headerRow}`);
  addMergeCell(doc, `K${headerRow}:M${headerRow}`);
  addMergeCell(doc, `B${subheaderRow}:C${subheaderRow}`);
  addMergeCell(doc, `B${totalRow}:C${totalRow}`);
  addDataValidation(doc, `G${subheaderRow}:J${subheaderRow}`, BCWDB_DROPDOWN_RANGE);
  addDataValidation(doc, `K${subheaderRow}:M${subheaderRow}`, NON_BCWDB_DROPDOWN_RANGE);

  return totalRow;

  function writeFundLabelIntoCell(doc2, cellEl, label) {
    const sharedIndex = SHARED[label];
    if (sharedIndex != null) {
      cellEl.setAttribute("t", "s");
      const v = el(doc2, "v");
      v.textContent = String(sharedIndex);
      cellEl.appendChild(v);
    } else {
      cellEl.setAttribute("t", "inlineStr");
      const is = el(doc2, "is");
      const t = el(doc2, "t");
      t.textContent = label;
      is.appendChild(t);
      cellEl.appendChild(is);
    }
  }
  function sumBy(arr) {
    return (arr || []).reduce((s, n) => s + (n || 0), 0);
  }

  function addMergeCell(doc2, ref) {
    let mergeCells = doc2.getElementsByTagNameNS(NS_MAIN, "mergeCells")[0];
    const worksheet = doc2.documentElement;
    if (!mergeCells) {
      mergeCells = el(doc2, "mergeCells");
      mergeCells.setAttribute("count", "0");
      // mergeCells must come after sheetData/dataValidations per schema order; append at end, fixed up later.
      worksheet.appendChild(mergeCells);
    }
    const mc = el(doc2, "mergeCell");
    mc.setAttribute("ref", ref);
    mergeCells.appendChild(mc);
    mergeCells.setAttribute("count", String(mergeCells.getElementsByTagNameNS(NS_MAIN, "mergeCell").length));
  }
  function addDataValidation(doc2, sqref, formula1) {
    let dvs = doc2.getElementsByTagNameNS(NS_MAIN, "dataValidations")[0];
    const worksheet = doc2.documentElement;
    if (!dvs) {
      dvs = el(doc2, "dataValidations");
      dvs.setAttribute("count", "0");
      worksheet.appendChild(dvs);
    }
    const dv = el(doc2, "dataValidation");
    dv.setAttribute("type", "list");
    dv.setAttribute("allowBlank", "1");
    dv.setAttribute("showInputMessage", "1");
    dv.setAttribute("showErrorMessage", "1");
    dv.setAttribute("sqref", sqref);
    const f1 = el(doc2, "formula1");
    f1.textContent = formula1;
    dv.appendChild(f1);
    dvs.appendChild(dv);
    dvs.setAttribute("count", String(dvs.getElementsByTagNameNS(NS_MAIN, "dataValidation").length));
  }
}

// The OOXML schema requires child elements of <worksheet> in a fixed order.
// Since appendBlock may have freshly appended <mergeCells>/<dataValidations>
// nodes at the very end (see addMergeCell/addDataValidation above), reorder
// them into a schema-valid sequence before serializing, or Excel will
// refuse to open the file / offer to repair it.
const WORKSHEET_CHILD_ORDER = [
  "sheetPr",
  "dimension",
  "sheetViews",
  "sheetFormatPr",
  "cols",
  "sheetData",
  "sheetCalcPr",
  "sheetProtection",
  "mergeCells",
  "conditionalFormatting",
  "dataValidations",
  "hyperlinks",
  "printOptions",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "drawing",
  "legacyDrawing",
];
function fixWorksheetChildOrder(doc) {
  const worksheet = doc.documentElement;
  const children = Array.from(worksheet.childNodes).filter((n) => n.nodeType === 1);
  children.sort((a, b) => {
    const ai = WORKSHEET_CHILD_ORDER.indexOf(a.localName);
    const bi = WORKSHEET_CHILD_ORDER.indexOf(b.localName);
    return (ai === -1 ? 999 : ai) - (bi === -1 ? 999 : bi);
  });
  children.forEach((c) => worksheet.appendChild(c));
}

// ---------------------------------------------------------------------------
// Zip / workbook-level plumbing
// ---------------------------------------------------------------------------

function sanitizeSheetName(name, fallback) {
  let n = (name || fallback || "Employee").replace(/[\\/*?:[\]]/g, " ").trim();
  if (!n) n = fallback || "Employee";
  return n.slice(0, 31);
}

function uniqueSheetName(desiredName, employeeId, usedNames) {
  let candidate = desiredName;
  if (!usedNames.has(candidate.toLowerCase())) return candidate;
  // Collision (e.g. two employees who share a display name) - disambiguate
  // with the Employee ID, trimming the base name so it still fits in 31 chars.
  const suffix = ` (${employeeId})`;
  candidate = `${desiredName.slice(0, 31 - suffix.length)}${suffix}`;
  let n = 1;
  while (usedNames.has(candidate.toLowerCase())) {
    const altSuffix = ` (${employeeId}-${n})`;
    candidate = `${desiredName.slice(0, 31 - altSuffix.length)}${altSuffix}`;
    n++;
  }
  return candidate;
}

async function loadWorkbookControlDocs(zip) {
  const workbookXml = await zip.file("xl/workbook.xml").async("string");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const ctXml = await zip.file("[Content_Types].xml").async("string");
  return {
    workbookDoc: parseXml(workbookXml),
    relsDoc: parseXml(relsXml),
    ctDoc: parseXml(ctXml),
  };
}

function nextSheetId(workbookDoc) {
  const sheets = Array.from(workbookDoc.getElementsByTagNameNS(NS_MAIN, "sheet"));
  const ids = sheets.map((s) => parseInt(s.getAttribute("sheetId"), 10) || 0);
  return (ids.length ? Math.max(...ids) : 0) + 1;
}
function nextRelId(relsDoc) {
  const rels = Array.from(relsDoc.getElementsByTagNameNS(NS_RELS, "Relationship"));
  const nums = rels
    .map((r) => r.getAttribute("Id"))
    .map((id) => (id && /^rId(\d+)$/.test(id) ? parseInt(id.match(/^rId(\d+)$/)[1], 10) : 0));
  return (nums.length ? Math.max(...nums) : 0) + 1;
}
function nextSheetFileNumber(zip) {
  const existing = Object.keys(zip.files).filter((p) => /^xl\/worksheets\/sheet\d+\.xml$/.test(p));
  const nums = existing.map((p) => parseInt(p.match(/sheet(\d+)\.xml$/)[1], 10));
  return (nums.length ? Math.max(...nums) : 0) + 1;
}

function registerNewSheet(ctx, sheetName) {
  const { workbookDoc, relsDoc, ctDoc, zip } = ctx;
  const sheetId = nextSheetId(workbookDoc);
  const rId = `rId${nextRelId(relsDoc)}`;
  const fileNum = nextSheetFileNumber(zip);
  const target = `worksheets/sheet${fileNum}.xml`;
  const partName = `/xl/${target}`;

  const sheetsEl = workbookDoc.getElementsByTagNameNS(NS_MAIN, "sheets")[0];
  const sheetEl = workbookDoc.createElementNS(NS_MAIN, "sheet");
  sheetEl.setAttribute("name", sheetName);
  sheetEl.setAttribute("sheetId", String(sheetId));
  sheetEl.setAttributeNS(NS_R, "r:id", rId);
  sheetsEl.appendChild(sheetEl);

  const relsRoot = relsDoc.documentElement;
  const relEl = relsDoc.createElementNS(NS_RELS, "Relationship");
  relEl.setAttribute("Id", rId);
  relEl.setAttribute(
    "Type",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
  );
  relEl.setAttribute("Target", target);
  relsRoot.appendChild(relEl);

  const ctRoot = ctDoc.documentElement;
  const overrideEl = ctDoc.createElementNS(NS_CT, "Override");
  overrideEl.setAttribute("PartName", partName);
  overrideEl.setAttribute(
    "ContentType",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
  );
  ctRoot.appendChild(overrideEl);

  return { sheetFilePath: `xl/${target}` };
}

function renameFirstSheet(workbookDoc, newName) {
  const sheetEl = workbookDoc.getElementsByTagNameNS(NS_MAIN, "sheet")[0];
  if (sheetEl) sheetEl.setAttribute("name", newName);
}

function stripPageSetupRel(doc) {
  // New sheets we create don't have a matching worksheets/_rels/sheetN.xml.rels
  // file, so drop the printerSettings relationship reference to avoid a
  // dangling-relationship repair prompt in Excel.
  const pageSetup = doc.getElementsByTagNameNS(NS_MAIN, "pageSetup")[0];
  if (pageSetup) pageSetup.removeAttributeNS(NS_R, "id");
}

function removeCalcChain(zip, ctDoc, relsDoc) {
  zip.remove("xl/calcChain.xml");
  const override = Array.from(ctDoc.getElementsByTagNameNS(NS_CT, "Override")).find(
    (o) => o.getAttribute("PartName") === "/xl/calcChain.xml"
  );
  if (override) override.parentNode.removeChild(override);
  const rel = Array.from(relsDoc.getElementsByTagNameNS(NS_RELS, "Relationship")).find((r) =>
    (r.getAttribute("Target") || "").includes("calcChain.xml")
  );
  if (rel) rel.parentNode.removeChild(rel);
}

function forceFullRecalcOnLoad(workbookDoc) {
  let calcPr = workbookDoc.getElementsByTagNameNS(NS_MAIN, "calcPr")[0];
  if (!calcPr) {
    calcPr = workbookDoc.createElementNS(NS_MAIN, "calcPr");
    workbookDoc.documentElement.appendChild(calcPr);
  }
  calcPr.setAttribute("fullCalcOnLoad", "1");
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

/**
 * @param {ArrayBuffer} templateArrayBuffer  - bundled blank Contractor_Timesheet_PY26.xlsx
 * @param {ArrayBuffer|null} ongoingArrayBuffer - the user's existing ongoing workbook, or null to start fresh
 * @param {Array} groupedEmployees - see lib/groupEntries.js for shape
 * @returns {{ blob: Blob, warnings: string[], summary: object[] }}
 */
export async function buildMergedWorkbook({ templateArrayBuffer, ongoingArrayBuffer, groupedEmployees }) {
  const warnings = [];
  const summary = [];

  // NOTE: loading the same ArrayBuffer into two separate JSZip instances at
  // once corrupts both (JSZip aliases the underlying buffer rather than
  // copying it). When we don't have an ongoing file, we reuse this same
  // `templateZip` instance as the output zip instead of loading it twice.
  const templateZip = await JSZip.loadAsync(templateArrayBuffer);
  const templateSheetXml = await templateZip.file("xl/worksheets/sheet1.xml").async("string");

  const outputZip = ongoingArrayBuffer ? await JSZip.loadAsync(ongoingArrayBuffer) : templateZip;

  const ctx = { zip: outputZip, ...(await loadWorkbookControlDocs(outputZip)) };

  // Map existing sheets (by their normalized EMP_ID marker) to file paths,
  // so repeat runs land on the SAME tab instead of creating a duplicate -
  // every pay period for an employee lives on that one tab, stacked as
  // sequential blocks below each other.
  const existingSheetByEmployeeId = new Map(); // employeeId -> sheetPath
  let isFreshWorkbook = !ongoingArrayBuffer;
  let usedTemplatesFirstSlot = false;

  const usedSheetNames = new Set(
    Array.from(ctx.workbookDoc.getElementsByTagNameNS(NS_MAIN, "sheet")).map((s) =>
      s.getAttribute("name").toLowerCase()
    )
  );

  if (ongoingArrayBuffer) {
    const sheetEls = Array.from(ctx.workbookDoc.getElementsByTagNameNS(NS_MAIN, "sheet"));
    for (const sheetEl of sheetEls) {
      const rId = sheetEl.getAttributeNS(NS_R, "id");
      const relEl = Array.from(ctx.relsDoc.getElementsByTagNameNS(NS_RELS, "Relationship")).find(
        (r) => r.getAttribute("Id") === rId
      );
      if (!relEl) continue;
      const target = relEl.getAttribute("Target").replace(/^\/?xl\//, "");
      const path = `xl/${target}`;
      const file = outputZip.file(path);
      if (!file) continue;
      const xmlText = await file.async("string");
      const doc = parseXml(xmlText);
      const sheetData = doc.getElementsByTagNameNS(NS_MAIN, "sheetData")[0];
      const empId = readEmployeeIdMarker(sheetData);
      if (empId) existingSheetByEmployeeId.set(empId, path);
    }
  }

  for (const employee of groupedEmployees) {
    let sheetPath;
    let doc;
    let sheetData;
    let isNewSheet = false;
    let assignedSheetName = null;

    if (existingSheetByEmployeeId.has(employee.employeeId)) {
      sheetPath = existingSheetByEmployeeId.get(employee.employeeId);
      const xmlText = await outputZip.file(sheetPath).async("string");
      doc = parseXml(xmlText);
      sheetData = doc.getElementsByTagNameNS(NS_MAIN, "sheetData")[0];
    } else if (isFreshWorkbook && !usedTemplatesFirstSlot) {
      // Reuse the workbook's original single sheet for the very first employee.
      doc = parseXml(templateSheetXml);
      sheetData = doc.getElementsByTagNameNS(NS_MAIN, "sheetData")[0];
      const sheetName = uniqueSheetName(
        sanitizeSheetName(employee.employeeName, employee.employeeId),
        employee.employeeId,
        usedSheetNames
      );
      usedSheetNames.add(sheetName.toLowerCase());
      const firstSheetEl = ctx.workbookDoc.getElementsByTagNameNS(NS_MAIN, "sheet")[0];
      const firstSheetRId = firstSheetEl.getAttributeNS(NS_R, "id");
      renameFirstSheet(ctx.workbookDoc, sheetName);
      const relEl = Array.from(ctx.relsDoc.getElementsByTagNameNS(NS_RELS, "Relationship")).find(
        (r) => r.getAttribute("Id") === firstSheetRId
      );
      const target = relEl.getAttribute("Target").replace(/^\/?xl\//, "");
      sheetPath = `xl/${target}`;
      usedTemplatesFirstSlot = true;
      isNewSheet = true;
      assignedSheetName = sheetName;
    } else {
      doc = parseXml(templateSheetXml);
      sheetData = doc.getElementsByTagNameNS(NS_MAIN, "sheetData")[0];
      const sheetName = uniqueSheetName(
        sanitizeSheetName(employee.employeeName, employee.employeeId),
        employee.employeeId,
        usedSheetNames
      );
      usedSheetNames.add(sheetName.toLowerCase());
      const { sheetFilePath } = registerNewSheet(ctx, sheetName);
      sheetPath = sheetFilePath;
      isNewSheet = true;
      assignedSheetName = sheetName;
    }

    if (isNewSheet) stripPageSetupRel(doc);

    const marker = readMarker(sheetData);
    const alreadyRecordedPeriods = new Set(marker ? marker.periods : []);
    const blockOneUsed = marker != null;

    const periodsWritten = [];
    let blockOneFilledThisRun = false;

    for (const period of employee.periods) {
      if (alreadyRecordedPeriods.has(period.key)) {
        warnings.push(
          `${employee.employeeName} (ID ${employee.employeeId}): pay period ${formatMDY(
            period.start
          )} - ${formatMDY(period.end)} was already on their tab, so it was skipped (not duplicated).`
        );
        continue;
      }

      const assignment = assignFundColumns(employee.employeeName, period, warnings);

      if (!blockOneUsed && !blockOneFilledThisRun) {
        fillBlock1(doc, sheetData, {
          employeeName: employee.employeeName,
          period,
          assignment,
          managerName: period.managerName,
        });
        blockOneFilledThisRun = true;
      } else {
        // Always compute where to append fresh - one row past whatever the
        // highest row number on the sheet currently is - rather than
        // trusting a stored "next row" value. That keeps this self-healing
        // even against a tab whose stored position was wrong (e.g. written
        // by a version of this tool that collided with the template's
        // signature rows), since it can never land on top of existing
        // content.
        const labelRow = maxRowNumber(sheetData) + 2;
        appendBlock(doc, sheetData, { period, assignment, labelRow });
      }
      alreadyRecordedPeriods.add(period.key);
      periodsWritten.push(period.key);
    }

    if (periodsWritten.length > 0) {
      writeMarker(doc, sheetData, employee.employeeId, {
        periods: Array.from(alreadyRecordedPeriods),
      });
      fixWorksheetChildOrder(doc);
      outputZip.file(sheetPath, serializeXml(doc));
      summary.push({
        employeeId: employee.employeeId,
        employeeName: employee.employeeName,
        periodsWritten: periodsWritten.length,
        sheetName: isNewSheet ? assignedSheetName : "(existing tab)",
      });
    }
  }

  removeCalcChain(outputZip, ctx.ctDoc, ctx.relsDoc);
  forceFullRecalcOnLoad(ctx.workbookDoc);

  outputZip.file("xl/workbook.xml", serializeXml(ctx.workbookDoc));
  outputZip.file("xl/_rels/workbook.xml.rels", serializeXml(ctx.relsDoc));
  outputZip.file("[Content_Types].xml", serializeXml(ctx.ctDoc));

  const blob = await outputZip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  return { blob, warnings, summary };
}

// ---------------------------------------------------------------------------
// Assigns each distinct fund seen in a pay period to one of the template's
// limited columns (4 BCWDB, 3 Non-BCWDB). If an employee genuinely worked
// more than 4 distinct BCWDB funds (or 3 Non-BCWDB) in a single 2-week
// period, the overflow is folded into the last available column and flagged
// - the template simply doesn't have room for more, so this needs a human
// to look at it.
// ---------------------------------------------------------------------------
function assignFundColumns(employeeName, period, warnings) {
  const bcwdb = period.bcwdbFunds.slice();
  const nonBcwdb = period.nonBcwdbFunds.slice();

  if (bcwdb.length > 4) {
    warnings.push(
      `${employeeName}: pay period ${formatMDY(period.start)} - ${formatMDY(
        period.end
      )} has ${bcwdb.length} different BCWDB-side funds, but the timesheet only has 4 columns for them. The extra fund(s) were combined into the last column - please review that tab.`
    );
    const overflow = bcwdb.splice(4);
    for (const fund of overflow) {
      fund.byDay.forEach((h, i) => {
        bcwdb[3].byDay[i] = (bcwdb[3].byDay[i] || 0) + (h || 0);
      });
      bcwdb[3].label = `${bcwdb[3].label} + ${fund.label}`;
    }
  }
  if (nonBcwdb.length > 3) {
    warnings.push(
      `${employeeName}: pay period ${formatMDY(period.start)} - ${formatMDY(
        period.end
      )} has ${nonBcwdb.length} different Non-BCWDB funds, but the timesheet only has 3 columns for them. The extra fund(s) were combined into the last column - please review that tab.`
    );
    const overflow = nonBcwdb.splice(3);
    for (const fund of overflow) {
      fund.byDay.forEach((h, i) => {
        nonBcwdb[2].byDay[i] = (nonBcwdb[2].byDay[i] || 0) + (h || 0);
      });
      nonBcwdb[2].label = `${nonBcwdb[2].label} + ${fund.label}`;
    }
  }

  return { bcwdb, nonBcwdb };
}
