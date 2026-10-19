"use client";

import { useState } from "react";
import { parseReportFile } from "@/lib/reportParser";
import { groupEntries } from "@/lib/groupEntries";
import { buildMergedWorkbook } from "@/lib/timesheetBuilder";
import { formatMDY } from "@/lib/payPeriods";

function StepBadge({ n, done }) {
  return (
    <span
      className={`inline-flex h-7 w-7 items-center justify-center rounded-full border-2 font-display text-sm font-bold ${
        done ? "border-teal bg-teal text-white" : "border-line text-ink-soft"
      }`}
    >
      {n}
    </span>
  );
}

function FileDropCard({ title, hint, file, onFile, accept, optional }) {
  const [dragOver, setDragOver] = useState(false);

  const handleFiles = (files) => {
    if (files && files[0]) onFile(files[0]);
  };

  return (
    <label
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        handleFiles(e.dataTransfer.files);
      }}
      className={`block cursor-pointer rounded-lg border-2 border-dashed p-6 transition-colors ${
        dragOver ? "border-teal bg-teal-tint" : "border-line bg-panel hover:border-teal/60"
      }`}
    >
      <input
        type="file"
        accept={accept}
        className="hidden"
        onChange={(e) => handleFiles(e.target.files)}
      />
      <div className="font-display text-sm font-semibold text-ink">
        {title}
        {optional && <span className="ml-2 font-sans text-xs font-normal text-ink-soft">(optional)</span>}
      </div>
      {hint && <div className="mt-1 text-xs text-ink-soft">{hint}</div>}
      {file ? (
        <div className="mt-3 inline-flex items-center gap-2 rounded-md bg-teal-tint px-3 py-1.5 text-sm font-medium text-teal-dark">
          <span className="font-mono">{file.name}</span>
        </div>
      ) : null}
    </label>
  );
}

function sumHours(arr) {
  return (arr || []).reduce((s, n) => s + (n || 0), 0);
}

/** Total hours across every pay period found for one grouped employee (see lib/groupEntries.js). */
function totalHoursForEmployee(employee) {
  let total = 0;
  for (const period of employee.periods) {
    total += sumHours(period.holidayByDay) + sumHours(period.ptoByDay);
    for (const fund of period.bcwdbFunds) total += sumHours(fund.byDay);
    for (const fund of period.nonBcwdbFunds) total += sumHours(fund.byDay);
  }
  return total;
}

export default function Home() {
  const [reportFile, setReportFile] = useState(null);
  const [ongoingFile, setOngoingFile] = useState(null);

  // idle | parsing | reviewing | processing | done | error
  const [status, setStatus] = useState("idle");
  const [errorMsg, setErrorMsg] = useState("");
  const [preMergeWarnings, setPreMergeWarnings] = useState([]); // parse + group warnings
  const [warnings, setWarnings] = useState([]);
  const [summary, setSummary] = useState([]);
  const [nonBcwdbMismatches, setNonBcwdbMismatches] = useState([]);
  const [nonBcwdbNotes, setNonBcwdbNotes] = useState([]);
  const [periodsFound, setPeriodsFound] = useState([]);
  const [groupedEmployees, setGroupedEmployees] = useState([]);
  const [selectedIds, setSelectedIds] = useState(new Set());
  const [downloadUrl, setDownloadUrl] = useState(null);
  const [downloadName, setDownloadName] = useState("");

  const canFindEmployees = !!reportFile && status !== "parsing" && status !== "processing";
  const canConfirmMerge = selectedIds.size > 0 && status !== "processing";

  function toggleEmployee(employeeId) {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(employeeId)) next.delete(employeeId);
      else next.add(employeeId);
      return next;
    });
  }

  function selectAll() {
    setSelectedIds(new Set(groupedEmployees.map((e) => e.employeeId)));
  }

  function selectNone() {
    setSelectedIds(new Set());
  }

  async function handleFindEmployees() {
    setStatus("parsing");
    setErrorMsg("");
    setWarnings([]);
    setSummary([]);
    setNonBcwdbMismatches([]);
    setNonBcwdbNotes([]);
    setDownloadUrl(null);

    try {
      const { entries, warnings: parseWarnings, reportedTotals } = await parseReportFile(reportFile);
      const { groupedEmployees: grouped, warnings: groupWarnings, periods } = groupEntries(entries, {
        reportedTotals,
      });
      setPeriodsFound(periods);
      setPreMergeWarnings([...parseWarnings, ...groupWarnings]);

      if (grouped.length === 0) {
        setWarnings([...parseWarnings, ...groupWarnings]);
        setStatus("error");
        setErrorMsg("No usable rows were found in the check report - nothing to merge.");
        return;
      }

      setGroupedEmployees(grouped);
      setSelectedIds(new Set(grouped.map((e) => e.employeeId)));
      setStatus("reviewing");
    } catch (err) {
      console.error(err);
      setStatus("error");
      setErrorMsg(err?.message || "Something went wrong while reading the report.");
    }
  }

  async function handleConfirmMerge() {
    setStatus("processing");
    setErrorMsg("");
    setWarnings([]);
    setSummary([]);
    setNonBcwdbMismatches([]);
    setNonBcwdbNotes([]);
    setDownloadUrl(null);

    try {
      const includedEmployees = groupedEmployees.filter((e) => selectedIds.has(e.employeeId));

      const templateArrayBuffer = await fetch("/template/Contractor_Timesheet_PY26.xlsx").then((r) =>
        r.arrayBuffer()
      );
      const ongoingArrayBuffer = ongoingFile ? await ongoingFile.arrayBuffer() : null;

      const {
        blob,
        warnings: buildWarnings,
        summary: buildSummary,
        nonBcwdbMismatches: buildMismatches,
        nonBcwdbNotes: buildNotes,
      } = await buildMergedWorkbook({
        templateArrayBuffer,
        ongoingArrayBuffer,
        groupedEmployees: includedEmployees,
        reportPeriods: periodsFound,
      });

      setWarnings([...preMergeWarnings, ...buildWarnings]);
      setSummary(buildSummary);
      setNonBcwdbMismatches(buildMismatches);
      setNonBcwdbNotes(buildNotes);

      const url = URL.createObjectURL(blob);
      setDownloadUrl(url);
      setDownloadName(
        ongoingFile ? ongoingFile.name.replace(/\.xlsx$/i, "") + " - updated.xlsx" : "Contractor Timesheets.xlsx"
      );
      setStatus("done");
    } catch (err) {
      console.error(err);
      setStatus("error");
      setErrorMsg(err?.message || "Something went wrong while processing the files.");
    }
  }

  return (
    <div className="min-h-screen bg-paper">
      <div className="mx-auto max-w-4xl px-6 py-12">
        <header className="mb-10">
          <h1 className="font-display text-3xl font-bold text-ink">Contractor Timesheet Merger</h1>
          <p className="mt-2 max-w-2xl text-ink-soft">
            Drop in a check report (PDF or Excel) and your ongoing Contractor Timesheet workbook. This
            tool matches each employee by ID, sorts their hours into the right fund columns, and gives
            you back an updated workbook - one tab per employee per pay period, ready to save.
            Everything runs in your browser; no file is ever uploaded anywhere.
          </p>
        </header>

        <section className="space-y-6 rounded-xl border border-line bg-panel p-6 shadow-sm">
          <div className="flex items-center gap-3">
            <StepBadge n={1} done={!!reportFile} />
            <h2 className="font-display text-lg font-semibold text-ink">Check report</h2>
          </div>
          <FileDropCard
            title="Drop the check report here (PDF or .xlsx)"
            hint="This is the report broken out by Project Code, e.g. Timesheets_3_30__4_26_26_Berks.pdf"
            file={reportFile}
            accept=".pdf,.xlsx,.xls"
            onFile={setReportFile}
          />

          <div className="flex items-center gap-3">
            <StepBadge n={2} done={true} />
            <h2 className="font-display text-lg font-semibold text-ink">Ongoing timesheet workbook</h2>
          </div>
          <FileDropCard
            title="Drop your ongoing Contractor Timesheet workbook here"
            hint="Leave this blank the very first time - the app will start a brand-new workbook from the blank template, one tab per employee."
            file={ongoingFile}
            accept=".xlsx"
            onFile={setOngoingFile}
            optional
          />

          <div className="flex items-center gap-3 pt-2">
            <StepBadge n={3} done={status === "reviewing" || status === "processing" || status === "done"} />
            <button
              onClick={handleFindEmployees}
              disabled={!canFindEmployees}
              className="rounded-md bg-teal px-5 py-2.5 font-display text-sm font-semibold text-white transition-colors hover:bg-teal-dark disabled:cursor-not-allowed disabled:opacity-40"
            >
              {status === "parsing" ? "Reading report..." : "Find employees"}
            </button>
          </div>
        </section>

        {status === "error" && errorMsg && (
          <div className="mt-6 rounded-lg border border-red-stamp/40 bg-red-tint p-4 text-sm text-red-stamp">
            {errorMsg}
          </div>
        )}

        {periodsFound.length > 0 && (
          <div className="mt-6 rounded-lg border border-line bg-panel p-4 text-sm text-ink-soft">
            <span className="font-semibold text-ink">Pay periods detected in this report: </span>
            {periodsFound.map((p) => `${formatMDY(p.start)} - ${formatMDY(p.end)}`).join(", ")}
          </div>
        )}

        {groupedEmployees.length > 0 && (
          <section className="mt-6 rounded-xl border border-line bg-panel p-6 shadow-sm">
            <div className="flex items-center justify-between gap-3">
              <h2 className="font-display text-lg font-semibold text-ink">
                Review employees ({selectedIds.size} of {groupedEmployees.length} selected)
              </h2>
              <div className="flex gap-2">
                <button
                  onClick={selectAll}
                  className="rounded-md border border-line px-3 py-1.5 text-xs font-semibold text-ink-soft transition-colors hover:border-teal hover:text-teal-dark"
                >
                  Select all
                </button>
                <button
                  onClick={selectNone}
                  className="rounded-md border border-line px-3 py-1.5 text-xs font-semibold text-ink-soft transition-colors hover:border-teal hover:text-teal-dark"
                >
                  Select none
                </button>
              </div>
            </div>
            <p className="mt-1 text-xs text-ink-soft">
              Unchecked employees are excluded from the merge. Uncheck anyone who shouldn&apos;t get a tab
              this run.
            </p>
            <table className="mt-4 w-full text-left text-sm">
              <thead>
                <tr className="border-b border-line text-ink-soft">
                  <th className="w-8 pb-2"></th>
                  <th className="pb-2 pr-4 font-medium">Employee</th>
                  <th className="pb-2 pr-4 font-medium">ID</th>
                  <th className="pb-2 font-medium">Total hours</th>
                </tr>
              </thead>
              <tbody>
                {groupedEmployees.map((employee) => (
                  <tr key={employee.employeeId} className="border-b border-line/60">
                    <td className="py-2">
                      <input
                        type="checkbox"
                        checked={selectedIds.has(employee.employeeId)}
                        onChange={() => toggleEmployee(employee.employeeId)}
                        className="h-4 w-4"
                      />
                    </td>
                    <td className="py-2 pr-4">{employee.employeeName}</td>
                    <td className="py-2 pr-4 font-mono text-ink-soft">{employee.employeeId}</td>
                    <td className="py-2">{totalHoursForEmployee(employee)}</td>
                  </tr>
                ))}
              </tbody>
            </table>

            <div className="mt-6 flex items-center gap-3">
              <button
                onClick={handleConfirmMerge}
                disabled={!canConfirmMerge}
                className="rounded-md bg-teal px-5 py-2.5 font-display text-sm font-semibold text-white transition-colors hover:bg-teal-dark disabled:cursor-not-allowed disabled:opacity-40"
              >
                {status === "processing" ? "Merging..." : "Merge into timesheet"}
              </button>
              {downloadUrl && (
                <a
                  href={downloadUrl}
                  download={downloadName}
                  className="rounded-md border-2 border-teal px-5 py-2.5 font-display text-sm font-semibold text-teal-dark transition-colors hover:bg-teal-tint"
                >
                  Download {downloadName}
                </a>
              )}
            </div>
          </section>
        )}

        {summary.length > 0 && (
          <section className="mt-6 rounded-xl border border-line bg-panel p-6 shadow-sm">
            <h2 className="font-display text-lg font-semibold text-ink">What got written</h2>
            <table className="mt-4 w-full text-left text-sm">
              <thead>
                <tr className="border-b border-line text-ink-soft">
                  <th className="pb-2 pr-4 font-medium">Employee</th>
                  <th className="pb-2 pr-4 font-medium">ID</th>
                  <th className="pb-2 pr-4 font-medium">Tab</th>
                  <th className="pb-2 font-medium">Pay periods added</th>
                </tr>
              </thead>
              <tbody>
                {summary.map((row, i) => (
                  <tr key={`${row.employeeId}-${i}`} className="border-b border-line/60">
                    <td className="py-2 pr-4">{row.employeeName}</td>
                    <td className="py-2 pr-4 font-mono text-ink-soft">{row.employeeId}</td>
                    <td className="py-2 pr-4">{row.sheetName}</td>
                    <td className="py-2">{row.periodsWritten}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </section>
        )}

        {nonBcwdbMismatches.length > 0 && (
          <section className="mt-6 rounded-xl border-2 border-yellow-400 bg-yellow-50 p-6 shadow-sm">
            <h2 className="font-display text-lg font-semibold text-ink">
              Non-BCWDB hours don&apos;t match the report
            </h2>
            <p className="mt-1 text-sm text-ink-soft">
              These employees&apos; Hours Worked Non-BCWDB were filled in as 8 minus their Berks hours
              each weekday, but the total doesn&apos;t match their report &quot;Total Hours&quot; minus
              their Berks hours. The values were still written - please check these tabs by hand.
            </p>
            <table className="mt-4 w-full text-left text-sm">
              <thead>
                <tr className="border-b border-yellow-400/60 text-ink-soft">
                  <th className="pb-2 pr-4 font-medium">Employee</th>
                  <th className="pb-2 pr-4 font-medium">ID</th>
                  <th className="pb-2 pr-4 font-medium">Expected hours</th>
                  <th className="pb-2 font-medium">Actual hours</th>
                </tr>
              </thead>
              <tbody>
                {nonBcwdbMismatches.map((m) => (
                  <tr key={m.employeeId} className="border-b border-yellow-400/30">
                    <td className="py-2 pr-4">{m.employeeName}</td>
                    <td className="py-2 pr-4 font-mono text-ink-soft">{m.employeeId}</td>
                    <td className="py-2 pr-4">{m.expectedHours}</td>
                    <td className="py-2">{m.actualHours}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </section>
        )}

        {nonBcwdbNotes.length > 0 && (
          <section className="mt-6 rounded-xl border border-line bg-gray-50 p-6 shadow-sm">
            <h2 className="font-display text-lg font-semibold text-ink">For your information</h2>
            <ul className="mt-3 space-y-2 text-sm text-ink-soft">
              {nonBcwdbNotes.map((note, i) => (
                <li key={i} className="flex gap-2">
                  <span aria-hidden>•</span>
                  <span>{note}</span>
                </li>
              ))}
            </ul>
          </section>
        )}

        {warnings.length > 0 && (
          <section className="mt-6 rounded-xl border border-amber/40 bg-amber-tint p-6 shadow-sm">
            <h2 className="font-display text-lg font-semibold text-ink">Please review</h2>
            <ul className="mt-3 space-y-2 text-sm text-ink">
              {warnings.map((w, i) => (
                <li key={i} className="flex gap-2">
                  <span aria-hidden>•</span>
                  <span>{w}</span>
                </li>
              ))}
            </ul>
          </section>
        )}
      </div>
    </div>
  );
}
