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
// Every employee+pay-period combination gets its own sheet/tab - always a
// fresh clone of the template's own built-in day-grid (rows 10-23),
// including its original formulas, so Excel keeps auto-calculating totals
// exactly like the blank template did.
//
// A small marker is stored in cells A1/A2 of every tab (invisible in normal
// use - column A isn't part of the printed/visible template) recording the
// employee ID and which pay period that tab covers. That's what makes it
// safe to run this tool again next pay period without duplicating anything:
// on the next run, every tab in the "ongoing" workbook is scanned for its
// marker so already-written employee+period combinations are skipped.
// Tabs written by older versions of this tool (which could hold more than
// one pay period on a single tab) are still recognized correctly, since
// their marker simply lists more than one period key.
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
// Filling the template's own block (rows 10-23) for one employee+pay-period
// tab. We deliberately leave all of the block's original formulas untouched
// - only the plain input cells get written.
// ---------------------------------------------------------------------------

function fillBlock1(doc, sheetData, { employeeId, employeeName, period, assignment, managerName, warnings }) {
  // Header fields.
  setExistingCellValue(doc, sheetData, "C", 7, { inlineText: employeeName }); // EMPLOYEE:
  setExistingCellValue(doc, sheetData, "C", 4, { number: toExcelSerial(period.start) }); // PAY PERIOD START DATE
  if (managerName) {
    setExistingCellValue(doc, sheetData, "C", 8, { inlineText: managerName }); // MANAGER:
  }

  writeFundHeaderRow(doc, sheetData, 11, assignment);
  writeDayRows(doc, sheetData, 12, period, assignment);
  refreshFormulaCachedValues(doc, sheetData, period, assignment);

  writeMarker(doc, sheetData, employeeId, { periods: [period.key] });
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

// The OOXML schema requires child elements of <worksheet> in a fixed order.
// Every tab we write starts from the template's own worksheet XML, so this
// is mostly a defensive no-op, but it's cheap insurance against Excel
// refusing to open the file / offering to repair it.
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

/** Short month-day tag for a pay period's start date, e.g. July 13 -> "7-13". */
function periodTag(period) {
  const d = period.start;
  return `${d.getMonth() + 1}-${d.getDate()}`;
}

/** Builds a tab name like "Lastname, First 7-13", truncating the name (never the period tag) to fit Excel's 31-char limit. */
function buildTabName(employeeName, employeeId, period) {
  const base = sanitizeSheetNameChars(employeeName, employeeId);
  const suffix = ` ${periodTag(period)}`;
  const maxBaseLen = Math.max(31 - suffix.length, 1);
  return `${base.slice(0, maxBaseLen)}${suffix}`;
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

  // Every employee+pay-period combination already written to some tab in
  // the ongoing workbook, so repeat runs skip periods instead of
  // duplicating them. Going forward each period gets its own tab, but a
  // single employee's periods can still be spread across several tabs (and
  // older workbooks may have several periods bundled onto one tab from
  // before this tool wrote one tab per period) - only the union matters.
  const alreadyWrittenPeriodsByEmployee = new Map(); // employeeId -> Set(periodKey)
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
      if (!empId) continue;
      const marker = readMarker(sheetData);
      const periods = marker && Array.isArray(marker.periods) ? marker.periods : [];
      if (!alreadyWrittenPeriodsByEmployee.has(empId)) {
        alreadyWrittenPeriodsByEmployee.set(empId, new Set());
      }
      const set = alreadyWrittenPeriodsByEmployee.get(empId);
      for (const key of periods) set.add(key);
    }
  }

  for (const employee of groupedEmployees) {
    const alreadyRecordedPeriods =
      alreadyWrittenPeriodsByEmployee.get(employee.employeeId) || new Set();

    for (const period of employee.periods) {
      if (alreadyRecordedPeriods.has(period.key)) {
        warnings.push(
          `${employee.employeeName} (ID ${employee.employeeId}): pay period ${formatMDY(
            period.start
          )} - ${formatMDY(period.end)} was already on their tab, so it was skipped (not duplicated).`
        );
        continue;
      }

      let doc, sheetData, sheetPath, assignedSheetName;
      const desiredName = buildTabName(employee.employeeName, employee.employeeId, period);

      if (isFreshWorkbook && !usedTemplatesFirstSlot) {
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
        const sheetName = uniqueSheetName(desiredName, employee.employeeId, usedSheetNames);
        usedSheetNames.add(sheetName.toLowerCase());
        const { sheetFilePath } = registerNewSheet(ctx, sheetName);
        sheetPath = sheetFilePath;
        assignedSheetName = sheetName;
      }

      stripPageSetupRel(doc);

      const assignment = assignFundColumns(employee.employeeName, period, warnings);
      fillBlock1(doc, sheetData, {
        employeeId: employee.employeeId,
        employeeName: employee.employeeName,
        period,
        assignment,
        managerName: period.managerName,
        warnings,
      });

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
