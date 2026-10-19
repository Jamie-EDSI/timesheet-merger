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
// Every employee + pay period combination gets its own standalone tab - a
// full, untouched clone of the template (header fields, all 10 day-rows
// with their original formulas so Excel keeps auto-calculating totals
// exactly like the blank template did, the totals/percentages row, AND its
// own EMPLOYEE SIGNATURE / MANAGER SIGNATURE lines). Nothing is stacked or
// appended into an existing tab, so there's no risk of a later period's
// rows colliding with an earlier period's content or with the template's
// own trailing rows.
//
// Tab naming: an employee's first/earliest pay period in the workbook gets
// the plain "Lastname, First" name (so anything referencing that name
// keeps resolving to real data); every pay period after that gets
// "Lastname, First (M-D)", using the pay period's start date to stay
// unique. Both are truncated to fit Excel's 31-character sheet name limit
// if needed - the date suffix is never the part that gets cut.
//
// A small marker is stored in cells A1/A2 of every tab (invisible in normal
// use - column A isn't part of the printed/visible template) recording the
// employee ID and the one pay period written to that tab. That's what makes
// it safe to run this tool again next pay period without duplicating
// anything: on the next run, every tab in the "ongoing" workbook is scanned
// for its marker, and matching is done by Employee ID with leading zeros
// stripped (some report exports zero-pad IDs and some don't - see
// lib/employeeId.js).
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
// Filling the template's own block (rows 10-23) for the one pay period a
// tab covers. We deliberately leave all of the block's original formulas
// untouched - only the plain input cells get written. The A1/A2 marker is
// written separately by the caller once this is done.
// ---------------------------------------------------------------------------

function fillBlock1(doc, sheetData, { employeeName, period, assignment, managerName, nonBcwdbFill }) {
  // Header fields.
  setExistingCellValue(doc, sheetData, "C", 7, { inlineText: employeeName }); // EMPLOYEE:
  // PAY PERIOD START DATE - a real date serial (the cell's own style already
  // has a date number format); C5 and C12:C21 chain off it.
  setExistingCellValue(doc, sheetData, "C", 4, { number: toExcelSerial(period.start) });
  if (managerName) {
    setExistingCellValue(doc, sheetData, "C", 8, { inlineText: managerName }); // MANAGER:
  }

  writeFundHeaderRow(doc, sheetData, 11, assignment);
  writeDayRows(doc, sheetData, 12, period, assignment);
  if (nonBcwdbFill) writeNonBcwdbFill(doc, sheetData, 12, nonBcwdbFill);
  guardPercentageRow(sheetData);
  const refreshed = refreshFormulaCachedValues(doc, sheetData, period, assignment, nonBcwdbFill);
  dropStaleCachedValues(sheetData, refreshed);
}

// Partially-allocated employees (see lib/groupEntries.js): the hours they
// worked outside Berks each weekday go in column K, with no fund label in
// K11 - it's "everything else", not a specific Non-BCWDB fund.
function writeNonBcwdbFill(doc, sheetData, startRow, fillByDay) {
  for (let i = 0; i < 10; i++) {
    if (fillByDay[i]) {
      setExistingCellValue(doc, sheetData, "K", startRow + i, { number: fillByDay[i] });
    }
  }
}

// G23:M23 divide by total non-leave hours (D22 - E22 - F22), which is 0 on
// an all-leave pay period. Wrap each in a zero check so Excel shows 0%
// instead of #DIV/0!. Only the formula text changes; styles stay as-is.
function guardPercentageRow(sheetData) {
  const denom = "($D$22-SUM($E$22:$F$22))";
  for (const col of ["G", "H", "I", "J", "K", "L", "M"]) {
    const c = findCell(findRow(sheetData, 23), col);
    const f = c && c.getElementsByTagNameNS(NS_MAIN, "f")[0];
    if (!f || f.textContent.startsWith("IF(")) continue;
    f.textContent = `IF(${denom}=0,0,${col}22/${denom})`;
  }
}

// Any formula cell we didn't compute a fresh cached value for would still
// show the blank template's number until Excel recalculates - drop its <v>
// so nothing stale is ever displayed (fullCalcOnLoad fills it back in).
function dropStaleCachedValues(sheetData, refreshedRefs) {
  for (const c of Array.from(sheetData.getElementsByTagNameNS(NS_MAIN, "c"))) {
    if (!c.getElementsByTagNameNS(NS_MAIN, "f")[0]) continue;
    if (refreshedRefs.has(c.getAttribute("r"))) continue;
    const v = c.getElementsByTagNameNS(NS_MAIN, "v")[0];
    if (v) c.removeChild(v);
    c.removeAttribute("t");
  }
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
function refreshFormulaCachedValues(doc, sheetData, period, assignment, nonBcwdbFill) {
  const refreshed = new Set(); // cell refs given a fresh cached value
  const setCached = (col, rowNum, value, opts) => {
    if (setFormulaCachedValue(sheetData, col, rowNum, value, opts)) refreshed.add(`${col}${rowNum}`);
  };
  const bcwdbCols = ["G", "H", "I", "J"];
  const nonBcwdbCols = ["K", "L", "M"];
  const fillOn = (i) => (nonBcwdbFill ? nonBcwdbFill[i] || 0 : 0);
  const dayTotal = (i) =>
    (period.holidayByDay[i] || 0) +
    (period.ptoByDay[i] || 0) +
    assignment.bcwdb.reduce((s, f) => s + (f.byDay[i] || 0), 0) +
    assignment.nonBcwdb.reduce((s, f) => s + (f.byDay[i] || 0), 0) +
    fillOn(i);
  const sumBy = (arr) => (arr || []).reduce((s, n) => s + (n || 0), 0);

  setCached("C", 5, toExcelSerial(period.end)); // =C4+11
  for (let i = 0; i < 10; i++) {
    const rowNum = 12 + i;
    setCached("C", rowNum, toExcelSerial(period.days[i]));
    setCached("D", rowNum, dayTotal(i));
  }

  const colTotal = (col) => {
    if (col === "D")
      return (
        sumBy(period.holidayByDay) +
        sumBy(period.ptoByDay) +
        assignment.bcwdb.reduce((s, f) => s + sumBy(f.byDay), 0) +
        assignment.nonBcwdb.reduce((s, f) => s + sumBy(f.byDay), 0) +
        sumBy(nonBcwdbFill)
      );
    if (col === "E") return sumBy(period.holidayByDay);
    if (col === "F") return sumBy(period.ptoByDay);
    if (col === "K" && nonBcwdbFill) return sumBy(nonBcwdbFill);
    const bi = bcwdbCols.indexOf(col);
    if (bi >= 0) return assignment.bcwdb[bi] ? sumBy(assignment.bcwdb[bi].byDay) : 0;
    const ni = nonBcwdbCols.indexOf(col);
    if (ni >= 0) return assignment.nonBcwdb[ni] ? sumBy(assignment.nonBcwdb[ni].byDay) : 0;
    return 0;
  };
  ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"].forEach((col) => {
    setCached(col, 22, colTotal(col));
  });

  // Row 23 percentages: hours in each fund column / total non-leave hours worked.
  const workedTotal = colTotal("D") - colTotal("E") - colTotal("F");
  ["G", "H", "I", "J", "K", "L", "M"].forEach((col) => {
    const pct = workedTotal > 0 ? colTotal(col) / workedTotal : 0;
    setCached(col, 23, pct, { clearErrorType: true });
  });
  return refreshed;
}

function setFormulaCachedValue(sheetData, colLetter, rowNum, value, { clearErrorType } = {}) {
  const row = findRow(sheetData, rowNum);
  const c = findCell(row, colLetter);
  if (!c) return false; // leave any cell we don't recognize alone
  if (clearErrorType) c.removeAttribute("t");
  let v = c.getElementsByTagNameNS(NS_MAIN, "v")[0];
  if (!v) {
    v = c.ownerDocument.createElementNS(NS_MAIN, "v");
    c.appendChild(v);
  }
  v.textContent = String(value);
  return true;
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

// The OOXML schema requires child elements of <worksheet> in a fixed order.
// Every tab we write starts from the template's own worksheet XML and only
// edits existing cells in place, so this is a defensive no-op in practice -
// but it's cheap insurance against Excel refusing to open the file / offering
// to repair it if that ever stops being true.
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

function sanitizeSheetNameChars(name, fallback) {
  let n = (name || fallback || "Employee").replace(/[\\/*?:[\]]/g, " ").trim();
  if (!n) n = fallback || "Employee";
  return n;
}

/** The plain "Lastname, First" tab name - used for an employee's first/earliest pay period. */
function sanitizeSheetName(name, fallback) {
  return sanitizeSheetNameChars(name, fallback).slice(0, 31);
}

/** Short month-day tag for a pay period's start date, e.g. July 13 -> "7-13". */
function periodTag(period) {
  const d = period.start;
  return `${d.getMonth() + 1}-${d.getDate()}`;
}

/** Builds a tab name like "Lastname, First (7-13)" for an employee's 2nd+ pay period, truncating the name (never the date suffix) to fit Excel's 31-char limit. */
function buildPeriodTabName(employeeName, employeeId, period) {
  const base = sanitizeSheetNameChars(employeeName, employeeId);
  const suffix = ` (${periodTag(period)})`;
  const maxBaseLen = Math.max(31 - suffix.length, 1);
  return `${base.slice(0, maxBaseLen)}${suffix}`;
}

/** True for tab names ending in a "(M-D)" pay-period suffix (see buildPeriodTabName), as opposed to a plain "Lastname, First" name or a "(EmployeeID)" collision-disambiguation suffix (which has no dash). */
function looksLikePeriodSuffixed(name) {
  return /\(\d{1,2}-\d{1,2}\)\s*$/.test(name || "");
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

// Strips a trailing " (EmployeeID)" or " (EmployeeID-2)" disambiguation
// suffix (see uniqueSheetName above) before comparing two sheet names, so a
// legacy tab and a freshly-desired name match even if one of them happens
// to carry that suffix.
function normalizeSheetNameForMatch(name) {
  return (name || "")
    .replace(/\s*\(\d+(?:-\d+)?\)\s*$/, "")
    .trim()
    .toLowerCase();
}

function renameSheetByPath(ctx, sheetPath, newName) {
  const relEl = Array.from(ctx.relsDoc.getElementsByTagNameNS(NS_RELS, "Relationship")).find(
    (r) => `xl/${r.getAttribute("Target").replace(/^\/?xl\//, "")}` === sheetPath
  );
  if (!relEl) return;
  const rId = relEl.getAttribute("Id");
  const sheetEl = Array.from(ctx.workbookDoc.getElementsByTagNameNS(NS_MAIN, "sheet")).find(
    (s) => s.getAttributeNS(NS_R, "id") === rId
  );
  if (sheetEl) sheetEl.setAttribute("name", newName);
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
 * @param {Array} reportPeriods - every pay period in the report's date range (groupEntries' `periods`)
 * @returns {{ blob: Blob, warnings: string[], summary: object[], nonBcwdbMismatches: object[], nonBcwdbNotes: string[] }}
 */
export async function buildMergedWorkbook({
  templateArrayBuffer,
  ongoingArrayBuffer,
  groupedEmployees,
  reportPeriods = [],
}) {
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

  // Every pay period already written to some tab for a given employee, so
  // repeat runs skip periods instead of duplicating them - the union across
  // however many of that employee's tabs exist, since each period now lives
  // on its own tab.
  const alreadyWrittenPeriodsByEmployee = new Map(); // employeeId -> Set(periodKey)
  // Employees who already have a tab under the plain "Lastname, First" name
  // (no period-date suffix) - that name is reserved for whichever pay period
  // claims it first, so later periods know to use the suffixed name instead.
  const employeesWithPlainTab = new Set(); // employeeId
  // Sheets from before this tool tagged tabs with an EMP_ID marker at all -
  // named just "Lastname, First", with no A1/A2 marker. Indexed by name so
  // a fresh merge for that same employee's first/earliest period can replace
  // the tab in place instead of leaving it behind alongside a newly-created,
  // differently-named one (which is what used to happen: the name collision
  // pushed the new data onto a suffixed tab, doubling up on the same person).
  const legacySheetPathByName = new Map(); // normalized name -> sheetPath
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
      if (!empId) {
        legacySheetPathByName.set(normalizeSheetNameForMatch(sheetEl.getAttribute("name")), path);
        continue;
      }
      const marker = readMarker(sheetData);
      const periods = marker && Array.isArray(marker.periods) ? marker.periods : [];
      if (!alreadyWrittenPeriodsByEmployee.has(empId)) {
        alreadyWrittenPeriodsByEmployee.set(empId, new Set());
      }
      const set = alreadyWrittenPeriodsByEmployee.get(empId);
      for (const key of periods) set.add(key);
      if (!looksLikePeriodSuffixed(sheetEl.getAttribute("name"))) {
        employeesWithPlainTab.add(empId);
      }
    }
  }

  // Reconcile each partially-allocated employee's Non-BCWDB fill BEFORE any
  // tab is written: the daily max(0, 8 - Berks hours) values, summed across
  // all their pay periods, should equal what the report's "Total Hours" line
  // implies they worked outside Berks. The daily rule can get that wrong
  // (e.g. someone who worked short days elsewhere), and a wrong number in
  // column K is worse than a blank one - so when it doesn't reconcile, K is
  // left blank on all their tabs and they're flagged for manual entry.
  //
  // Part of a shortfall can be explained by report pay periods where the
  // employee has no Berks rows at all - no tab is created for those, so
  // their 8 hours x 10 weekdays never get written anywhere. When that
  // accounts for the whole difference, the fill is still written and it's
  // an informational note, not a warning.
  const nonBcwdbMismatches = [];
  const nonBcwdbNotes = [];
  const employeesWithoutFill = new Set(); // employeeId - reconciliation failed, K stays blank
  for (const employee of groupedEmployees) {
    if (!employee.isPartiallyAllocated) continue;
    const expectedHours = employee.reportedTotalHours - employee.berksTotalHours;
    const actualHours = employee.periods.reduce(
      (s, p) => s + (p.nonBcwdbFillByDay || []).reduce((t, h) => t + h, 0),
      0
    );
    const ownPeriodKeys = new Set(employee.periods.map((p) => p.key));
    const periodsWithoutTab = reportPeriods.filter((p) => !ownPeriodKeys.has(p.key));
    const explainedHours = periodsWithoutTab.reduce((s, p) => s + 8 * p.days.length, 0);
    if (Math.abs(expectedHours - actualHours) < 1e-9) continue;
    if (periodsWithoutTab.length && Math.abs(actualHours + explainedHours - expectedHours) < 1e-9) {
      const ranges = periodsWithoutTab.map((p) => `${formatMD(p.start)}-${formatMD(p.end)}`).join(", ");
      nonBcwdbNotes.push(
        `${employee.employeeName}: no Berks hours in ${ranges}, so no tab was created for ${
          periodsWithoutTab.length === 1 ? "that period" : "those periods"
        } (${explainedHours} non-Berks hours not shown).`
      );
      continue;
    }
    employeesWithoutFill.add(employee.employeeId);
    nonBcwdbMismatches.push({
      employeeId: employee.employeeId,
      employeeName: employee.employeeName,
      expectedHours,
      actualHours,
    });
  }

  for (const employee of groupedEmployees) {
    const alreadyRecordedPeriods =
      alreadyWrittenPeriodsByEmployee.get(employee.employeeId) || new Set();
    let employeeHasPlainTab = employeesWithPlainTab.has(employee.employeeId);

    for (const period of employee.periods) {
      if (alreadyRecordedPeriods.has(period.key)) {
        warnings.push(
          `${employee.employeeName} (ID ${employee.employeeId}): pay period ${formatMDY(
            period.start
          )} - ${formatMDY(period.end)} was already on their tab, so it was skipped (not duplicated).`
        );
        continue;
      }

      // The plain "Lastname, First" name is reserved for this employee's
      // first/earliest pay period (employee.periods is sorted ascending);
      // every period after that gets the "(M-D)" suffix.
      const usesPlainName = !employeeHasPlainTab;
      if (usesPlainName) employeeHasPlainTab = true;

      const plainName = sanitizeSheetName(employee.employeeName, employee.employeeId);
      const desiredName = usesPlainName ? plainName : buildPeriodTabName(employee.employeeName, employee.employeeId, period);

      const legacyMatchKey = usesPlainName ? normalizeSheetNameForMatch(plainName) : null;
      const legacySheetPath = legacyMatchKey ? legacySheetPathByName.get(legacyMatchKey) : null;

      let sheetPath;
      let doc;
      let sheetData;
      let assignedSheetName;

      if (legacySheetPath) {
        // A pre-migration tab already exists under this employee's plain
        // name with no EMP_ID marker at all. It's safe to fully replace:
        // start from a clean template clone (not the legacy tab's old,
        // possibly incomplete/broken content) and write it back to that
        // same sheet path, so the workbook ends up with one correctly-
        // tracked tab instead of the old broken one plus a new duplicate.
        legacySheetPathByName.delete(legacyMatchKey); // don't let a 2nd employee reuse the same legacy tab
        sheetPath = legacySheetPath;
        doc = parseXml(templateSheetXml);
        sheetData = doc.getElementsByTagNameNS(NS_MAIN, "sheetData")[0];
        renameSheetByPath(ctx, sheetPath, plainName);
        usedSheetNames.add(plainName.toLowerCase());
        assignedSheetName = plainName;
      } else if (isFreshWorkbook && !usedTemplatesFirstSlot) {
        // Reuse the workbook's original single sheet for the very first tab.
        doc = parseXml(templateSheetXml);
        sheetData = doc.getElementsByTagNameNS(NS_MAIN, "sheetData")[0];
        const sheetName = uniqueSheetName(desiredName, employee.employeeId, usedSheetNames);
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
        assignedSheetName = sheetName;
      } else {
        doc = parseXml(templateSheetXml);
        sheetData = doc.getElementsByTagNameNS(NS_MAIN, "sheetData")[0];
        // Genuine collision: a DIFFERENT employee already has a tab under
        // this same display name (not a stale legacy tab, since those are
        // matched and consumed above) - disambiguate with an ID suffix.
        const sheetName = uniqueSheetName(desiredName, employee.employeeId, usedSheetNames);
        usedSheetNames.add(sheetName.toLowerCase());
        const { sheetFilePath } = registerNewSheet(ctx, sheetName);
        sheetPath = sheetFilePath;
        assignedSheetName = sheetName;
      }

      stripPageSetupRel(doc);

      const assignment = assignFundColumns(employee.employeeName, period, warnings);
      let nonBcwdbFill = employeesWithoutFill.has(employee.employeeId)
        ? null
        : period.nonBcwdbFillByDay || null;
      if (nonBcwdbFill && assignment.nonBcwdb.length > 0) {
        // Column K already holds a real Non-BCWDB fund from the report.
        warnings.push(
          `${employee.employeeName} (ID ${employee.employeeId}): pay period ${formatMDY(
            period.start
          )} - ${formatMDY(period.end)} already has a Non-BCWDB fund in column K, so the non-Berks hours from the report's "Total Hours" line were NOT filled in - please add them by hand.`
        );
        nonBcwdbFill = null;
      }
      fillBlock1(doc, sheetData, {
        employeeName: employee.employeeName,
        period,
        assignment,
        managerName: period.managerName,
        nonBcwdbFill,
      });
      writeMarker(doc, sheetData, employee.employeeId, { periods: [period.key] });

      fixWorksheetChildOrder(doc);
      outputZip.file(sheetPath, serializeXml(doc));

      summary.push({
        employeeId: employee.employeeId,
        employeeName: employee.employeeName,
        periodsWritten: 1,
        sheetName: assignedSheetName,
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

  return { blob, warnings, summary, nonBcwdbMismatches, nonBcwdbNotes };
}

/** "4/13" - month/day only, for compact period ranges in notes. */
function formatMD(date) {
  return `${date.getMonth() + 1}/${date.getDate()}`;
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
