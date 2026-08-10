import * as XLSX from "xlsx";

// ---------------------------------------------------------------------------
// Parses a "check report" file - either the PDF export or an Excel export of
// the same report - into a flat list of normalized rows:
//
//   { employeeId, employeeName, date: Date, projectCode: "525002", hours: 8 }
//
// The report's real layout (seen in the sample PDF) repeats, per employee:
//   header row, then daily detail rows like:
//     "<ID> <Name>, <First> <Title> <Location> <M/D/YYYY> REG SALARY <code> Approved [Approver] <qty>"
//   followed by a "SUMMARY - TOTAL HOURS BY PROJECT" block and grand totals,
//   which we skip.
// ---------------------------------------------------------------------------

// Anchors on: employee ID, a date (M/D/YYYY), a 6-digit project code, the
// literal "Approved" status, an optional approver name, then the hours
// (always the last number on the line). This is deliberately anchored on
// the parts of the line that never move, rather than fixed column widths,
// since PDF text extraction doesn't preserve exact spacing.
const LINE_RE =
  /(\d{2,10})\s+(.+?)\s+(\d{1,2}\/\d{1,2}\/\d{4})\s+(?:REG SALARY\s+)?(\d{6})\s+Approved\s*(.*?)\s*(\d+(?:\.\d+)?)\s*$/gm;

export function parseReportText(text) {
  const entries = [];
  const warnings = [];
  let match;
  LINE_RE.lastIndex = 0;
  while ((match = LINE_RE.exec(text)) !== null) {
    const [, idRaw, nameBlob, dateRaw, code, , qtyRaw] = match;
    const date = parseMDY(dateRaw);
    const hours = parseFloat(qtyRaw);
    if (!date || Number.isNaN(hours)) continue;
    entries.push({
      employeeId: idRaw.trim(),
      employeeName: extractDisplayName(nameBlob),
      date,
      projectCode: code,
      hours,
    });
  }
  if (entries.length === 0) {
    warnings.push(
      "Couldn't find any recognizable report rows in this file. Double-check it's a check report in the expected format."
    );
  }
  return { entries, warnings };
}

/** Best-effort "Lastname, Firstname" extraction from the blob of name+title+location text. Falls back to the raw blob if it doesn't look like "Last, First". */
function extractDisplayName(blob) {
  const m = blob.match(/^([^,]+,\s*\S+)/);
  return m ? m[1].trim() : blob.trim();
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
      const dateCol = row.findIndex((c) => c === "date");
      const codeCol = row.findIndex((c) => c.includes("project code"));
      const qtyCol = row.findIndex((c) => c.includes("reported quantity"));
      const nameCol = row.findIndex((c) => c.includes("name in reporting"));
      if (idCol >= 0 && dateCol >= 0 && codeCol >= 0 && qtyCol >= 0) {
        headerIdx = i;
        colIndex = { idCol, dateCol, codeCol, qtyCol, nameCol };
        break;
      }
    }

    if (headerIdx === -1) {
      // Fall back to treating this sheet as a raw text blob and reuse the
      // PDF-style line parser - handles exports that don't have a clean
      // tabular header row.
      const asText = rows.map((r) => r.join(" ")).join("\n");
      const { entries: fallbackEntries } = parseReportText(asText);
      entries.push(...fallbackEntries);
      continue;
    }

    for (let i = headerIdx + 1; i < rows.length; i++) {
      const row = rows[i];
      const idVal = String(row[colIndex.idCol] ?? "").trim();
      const dateVal = row[colIndex.dateCol];
      const codeVal = String(row[colIndex.codeCol] ?? "").trim();
      const qtyVal = row[colIndex.qtyCol];

      if (!/^\d{2,10}$/.test(idVal)) continue; // not a data row (blank/summary/header repeat)
      if (!/^\d{6}$/.test(codeVal)) continue;
      const hours = parseFloat(qtyVal);
      if (Number.isNaN(hours)) continue;

      const date = dateVal instanceof Date ? dateVal : parseMDY(String(dateVal));
      if (!date) continue;

      const nameRaw =
        colIndex.nameCol >= 0 ? String(row[colIndex.nameCol] ?? "").trim() : "";

      entries.push({
        employeeId: idVal,
        employeeName: extractDisplayName(nameRaw || idVal),
        date,
        projectCode: codeVal,
        hours,
      });
    }
  }

  if (entries.length === 0) {
    warnings.push(
      "Couldn't find any recognizable report rows in this Excel file. Double-check it's a check report in the expected format."
    );
  }

  return { entries, warnings };
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
