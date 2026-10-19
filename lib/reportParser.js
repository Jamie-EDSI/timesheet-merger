import * as XLSX from "xlsx";
import { normalizeEmployeeId } from "./employeeId.js";

// ---------------------------------------------------------------------------
// Parses a "check report" file - either the PDF export or an Excel export of
// the same report - into a flat list of normalized rows:
//
//   { employeeId, employeeName, date: Date, projectCode: "525002", payCode,
//     approver, hours: 8 }
//
// The report's real layout (seen in the sample PDF) repeats, per employee:
//   header row, then daily detail rows like:
//     "<ID> <Name>, <First> <Title> <Location> <M/D/YYYY> [REG SALARY|HOLIDAY] <code> Approved [Approver] <qty>"
//   followed by a "SUMMARY - TOTAL HOURS BY PROJECT" block and grand totals,
//   which we skip. The pay code between the date and project code is
//   optional - some rows have none at all.
// ---------------------------------------------------------------------------

// The return value also carries `reportedTotals`: { [employeeId]: hours }.
// A few employees' blocks have an extra line right after their GRAND TOTAL,
// e.g. "160 Total Hours" - the hours they actually worked across ALL
// programs, not just Berks. lib/groupEntries.js uses it to fill in the
// "Hours Worked Non-BCWDB" column for those partially-allocated employees.
// ---------------------------------------------------------------------------

// Anchors on: employee ID, a date (M/D/YYYY), an optional pay code, a
// 6-digit project code, the literal "Approved" status, an optional approver
// name, then the hours (always the last number on the line). This is
// deliberately anchored on the parts of the line that never move, rather
// than fixed column widths, since PDF text extraction doesn't preserve
// exact spacing.
const LINE_RE =
  /(\d{2,10})\s+(.+?)\s+(\d{1,2}\/\d{1,2}\/\d{4})\s+(?:(REG SALARY|HOLIDAY)\s+)?(\d{6})\s+Approved\s*(.*?)\s*(\d+(?:\.\d+)?)\s*$/gm;

// "GRAND TOTAL <n>" followed by "<m> Total Hours". Text extraction sometimes
// glues the next block's header straight onto the end ("...Total HoursEmployee
// ID..."), so nothing is required after "Total Hours". Requiring a number
// before "Total Hours" keeps this from matching the "Project Code   Total
// Hours" column header.
const REPORTED_TOTAL_RE = /GRAND TOTAL\s+\d+(?:\.\d+)?\s+(\d+(?:\.\d+)?)\s*Total Hours/gi;

export function parseReportText(text) {
  const entries = [];
  const warnings = [];
  const entryPositions = []; // [textIndex, employeeId] for each detail row, in order
  let match;
  LINE_RE.lastIndex = 0;
  while ((match = LINE_RE.exec(text)) !== null) {
    const [, idRaw, nameBlob, dateRaw, payCode, code, approverRaw, qtyRaw] = match;
    const date = parseMDY(dateRaw);
    const hours = parseFloat(qtyRaw);
    if (!date || Number.isNaN(hours)) continue;
    const employeeId = normalizeEmployeeId(idRaw);
    entryPositions.push([match.index, employeeId]);
    entries.push({
      employeeId,
      employeeName: extractDisplayName(nameBlob, { employeeId, warnings }),
      date,
      projectCode: code,
      payCode: payCode || null,
      approver: cleanApproverName(approverRaw),
      hours,
    });
  }

  // Each "Total Hours" line belongs to the employee whose detail rows came
  // right before that GRAND TOTAL.
  const reportedTotals = {};
  REPORTED_TOTAL_RE.lastIndex = 0;
  while ((match = REPORTED_TOTAL_RE.exec(text)) !== null) {
    let owner = null;
    for (const [pos, employeeId] of entryPositions) {
      if (pos > match.index) break;
      owner = employeeId;
    }
    if (owner) reportedTotals[owner] = parseFloat(match[1]);
  }

  if (entries.length === 0) {
    warnings.push(
      "Couldn't find any recognizable report rows in this file. Double-check it's a check report in the expected format."
    );
  }
  return { entries, warnings, reportedTotals };
}

/** Strips a trailing "(Terminated)" suffix (however it's cased/spaced) off an approver name. */
function cleanApproverName(raw) {
  return (raw || "").replace(/\s*\(Terminated\)\s*$/i, "").trim();
}

// Matches a numeric Employee ID glued onto the front of a Name field, e.g.
// "1125662   Wilson, Amanda" - some report exports concatenate the ID and
// name into a single field instead of keeping them as separate columns.
const LEADING_ID_RE = /^\d{2,10}[\s,]+(?=\S)/;

/** Best-effort "Lastname, Firstname" extraction from the blob of name+title+location text. Falls back to the raw blob if it doesn't look like "Last, First". Strips (and flags) a leading Employee ID glued onto the front, if present. */
function extractDisplayName(blob, { employeeId, warnings } = {}) {
  let working = blob;
  const idMatch = blob.match(LEADING_ID_RE);
  if (idMatch) {
    working = blob.slice(idMatch[0].length);
    if (warnings) {
      warnings.push(
        `Employee ID ${employeeId}: the Name field ("${blob.trim()}") looked like it had an Employee ID glued onto the front, which can break employee matching and create a duplicate tab. Stripped it down to "${working.trim()}" - please double-check this row.`
      );
    }
  }
  const m = working.match(/^([^,]+,\s*\S+)/);
  return m ? m[1].trim() : working.trim();
}

function parseMDY(str) {
  const m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  const [, mo, da, yr] = m;
  return new Date(Number(yr), Number(mo) - 1, Number(da));
}

export async function parsePdfReport(arrayBuffer) {
  const pdfjsLib = await import("pdfjs-dist/build/pdf.mjs");
  pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
    "pdfjs-dist/build/pdf.worker.mjs",
    import.meta.url
  ).toString();

  const doc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let fullText = "";
  for (let pageNum = 1; pageNum <= doc.numPages; pageNum++) {
    const page = await doc.getPage(pageNum);
    const content = await page.getTextContent();

    // Group text items into lines by their y-position so that words on the
    // same visual line end up in the same text line, even though pdf.js
    // hands them back as separate positioned items.
    const lines = new Map(); // roundedY -> [{x, str}]
    for (const item of content.items) {
      const y = Math.round(item.transform[5]);
      if (!lines.has(y)) lines.set(y, []);
      lines.get(y).push({ x: item.transform[4], str: item.str });
    }
    const sortedY = Array.from(lines.keys()).sort((a, b) => b - a); // top to bottom
    for (const y of sortedY) {
      const words = lines.get(y).sort((a, b) => a.x - b.x);
      fullText += words.map((w) => w.str).join(" ") + "\n";
    }
  }
  return parseReportText(fullText);
}

export function parseExcelReport(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const entries = [];
  const warnings = [];
  const reportedTotals = {};

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
    if (!rows.length) continue;

    // Find the header row and the columns we care about by name.
    let headerIdx = -1;
    let colIndex = {};
    for (let i = 0; i < Math.min(rows.length, 20); i++) {
      const row = rows[i].map((c) => String(c).trim().toLowerCase());
      const idCol = row.findIndex((c) => c.includes("employee id"));
      const dateCol = row.findIndex((c) => c === "date" || c === "reported date");
      const codeCol = row.findIndex((c) => c.includes("project code"));
      const qtyCol = row.findIndex((c) => c.includes("reported quantity"));
      const nameCol = row.findIndex((c) => c.includes("name in reporting"));
      const payCodeCol = row.findIndex((c) => c.includes("pay code"));
      const approverCol = row.findIndex((c) => c.includes("approver"));
      if (idCol >= 0 && dateCol >= 0 && codeCol >= 0 && qtyCol >= 0) {
        headerIdx = i;
        colIndex = { idCol, dateCol, codeCol, qtyCol, nameCol, payCodeCol, approverCol };
        break;
      }
    }

    if (headerIdx === -1) {
      // Fall back to treating this sheet as a raw text blob and reuse the
      // PDF-style line parser - handles exports that don't have a clean
      // tabular header row.
      const asText = rows.map((r) => r.join(" ")).join("\n");
      const { entries: fallbackEntries, reportedTotals: fallbackTotals } = parseReportText(asText);
      entries.push(...fallbackEntries);
      Object.assign(reportedTotals, fallbackTotals);
      continue;
    }

    let lastEmployeeId = null;
    for (let i = headerIdx + 1; i < rows.length; i++) {
      const row = rows[i];
      const idVal = String(row[colIndex.idCol] ?? "").trim();
      const dateVal = row[colIndex.dateCol];
      const codeVal = String(row[colIndex.codeCol] ?? "").trim();
      const qtyVal = row[colIndex.qtyCol];

      // A "160 Total Hours" line (in one cell or split across two) after an
      // employee's GRAND TOTAL - belongs to the last employee whose rows we saw.
      const totalMatch = row
        .map((c) => String(c).trim())
        .filter(Boolean)
        .join(" ")
        .match(/^(\d+(?:\.\d+)?)\s*Total Hours$/i);
      if (totalMatch) {
        if (lastEmployeeId) reportedTotals[lastEmployeeId] = parseFloat(totalMatch[1]);
        continue;
      }

      if (!/^\d{2,10}$/.test(idVal)) continue; // not a data row (blank/summary/header repeat)
      if (!/^\d{6}$/.test(codeVal)) continue;
      const hours = parseFloat(qtyVal);
      if (Number.isNaN(hours)) continue;

      const date = dateVal instanceof Date ? dateVal : parseMDY(String(dateVal));
      if (!date) continue;

      const nameRaw =
        colIndex.nameCol >= 0 ? String(row[colIndex.nameCol] ?? "").trim() : "";
      const payCodeRaw =
        colIndex.payCodeCol >= 0 ? String(row[colIndex.payCodeCol] ?? "").trim() : "";
      const approverRaw =
        colIndex.approverCol >= 0 ? String(row[colIndex.approverCol] ?? "").trim() : "";

      const employeeId = normalizeEmployeeId(idVal);
      lastEmployeeId = employeeId;
      entries.push({
        employeeId,
        employeeName: extractDisplayName(nameRaw || idVal, { employeeId, warnings }),
        date,
        projectCode: codeVal,
        payCode: payCodeRaw || null,
        approver: cleanApproverName(approverRaw),
        hours,
      });
    }
  }

  if (entries.length === 0) {
    warnings.push(
      "Couldn't find any recognizable report rows in this Excel file. Double-check it's a check report in the expected format."
    );
  }

  return { entries, warnings, reportedTotals };
}

export async function parseReportFile(file) {
  const buf = await file.arrayBuffer();
  const isPdf =
    file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf");
  if (isPdf) {
    return parsePdfReport(buf);
  }
  return parseExcelReport(buf);
}
