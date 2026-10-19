import { computePayPeriods, dateKey } from "./payPeriods.js";
import { lookupProjectCode } from "./fundMap.js";

const IGNORED_APPROVER_RE = /^automatically approved$/i;

/**
 * Takes the flat list of parsed report rows and turns it into the structure
 * lib/timesheetBuilder.js needs: one entry per employee, each with a list
 * of pay periods, each period holding per-day hours already sorted into
 * "holiday", "pto", "bcwdb funds", and "non-bcwdb funds".
 *
 * `reportedTotals` ({ [employeeId]: hours }, from lib/reportParser.js) is
 * each employee's "Total Hours" line, if the report had one. When it's more
 * than their Berks total, the employee is only partially allocated to Berks
 * and each weekday gets `nonBcwdbFillByDay` = max(0, 8 - that day's Berks
 * hours), which the builder writes into the HOURS WORKED NON-BCWDB column.
 */
export function groupEntries(entries, { reportedTotals = {} } = {}) {
  const warnings = [];
  const unmappedCodesSeen = new Map(); // code -> {label, note}

  const allDates = entries.map((e) => e.date);
  const { periods, periodOf } = computePayPeriods(allDates);
  const periodByKey = new Map(periods.map((p) => [dateKey(p.start), p]));

  // employeeId -> { name, periods: Map(periodKey -> workingPeriod) }
  const employees = new Map();

  for (const entry of entries) {
    const isHolidayPayCode = (entry.payCode || "").toUpperCase() === "HOLIDAY";
    const fund = lookupProjectCode(entry.projectCode);
    if (!isHolidayPayCode && fund.category === "unmapped" && !unmappedCodesSeen.has(entry.projectCode)) {
      unmappedCodesSeen.set(entry.projectCode, fund);
    }

    if (!employees.has(entry.employeeId)) {
      employees.set(entry.employeeId, { name: entry.employeeName, periods: new Map(), berksTotalHours: 0 });
    }
    const employee = employees.get(entry.employeeId);
    employee.berksTotalHours += entry.hours; // every project code, leave and holiday included
    if (entry.employeeName && entry.employeeName.length > employee.name.length) {
      // Prefer the fuller-looking name if we saw a better one later.
      employee.name = entry.employeeName;
    }

    const periodKey = periodOf(entry.date);
    const periodTemplate = periodByKey.get(periodKey);
    if (!employee.periods.has(periodKey)) {
      employee.periods.set(periodKey, {
        key: periodKey,
        start: periodTemplate.start,
        end: periodTemplate.end,
        days: periodTemplate.days,
        holidayByDay: new Array(10).fill(0),
        ptoByDay: new Array(10).fill(0),
        bcwdbFundsByLabel: new Map(), // label -> { label, matchesDropdown, byDay: number[10] }
        nonBcwdbFundsByLabel: new Map(),
        approverCounts: new Map(), // approver name -> count
      });
    }
    const period = employee.periods.get(periodKey);

    if (entry.approver && !IGNORED_APPROVER_RE.test(entry.approver)) {
      period.approverCounts.set(entry.approver, (period.approverCounts.get(entry.approver) || 0) + 1);
    }

    const dayIndex = period.days.findIndex((d) => dateKey(d) === dateKey(entry.date));
    if (dayIndex === -1) {
      warnings.push(
        `${employee.name} (ID ${entry.employeeId}): a row dated ${entry.date.toDateString()} didn't line up with a weekday in its pay period (maybe a weekend date in the report?) and was skipped.`
      );
      continue;
    }

    if (isHolidayPayCode) {
      period.holidayByDay[dayIndex] += entry.hours;
    } else if (fund.category === "leave") {
      const target = fund.leaveColumn === "holiday" ? period.holidayByDay : period.ptoByDay;
      target[dayIndex] += entry.hours;
    } else {
      const bucket = fund.category === "nonBcwdb" ? period.nonBcwdbFundsByLabel : period.bcwdbFundsByLabel;
      if (!bucket.has(fund.label)) {
        bucket.set(fund.label, { label: fund.label, matchesDropdown: fund.matchesDropdown, byDay: new Array(10).fill(0) });
      }
      bucket.get(fund.label).byDay[dayIndex] += entry.hours;
    }
  }

  for (const [code, fund] of unmappedCodesSeen) {
    warnings.push(`Project Code ${code} ("${fund.label}"): ${fund.note}`);
  }

  const groupedEmployees = Array.from(employees.entries()).map(([employeeId, employee]) => {
    const reportedTotalHours = reportedTotals[employeeId] ?? null;
    const isPartiallyAllocated =
      reportedTotalHours != null && reportedTotalHours > employee.berksTotalHours;
    return {
      employeeId,
      employeeName: employee.name,
      berksTotalHours: employee.berksTotalHours,
      reportedTotalHours,
      isPartiallyAllocated,
      periods: Array.from(employee.periods.values())
        .map((p) => {
          const bcwdbFunds = Array.from(p.bcwdbFundsByLabel.values());
          const nonBcwdbFunds = Array.from(p.nonBcwdbFundsByLabel.values());
          return {
            ...p,
            bcwdbFunds,
            nonBcwdbFunds,
            managerName: mostFrequent(p.approverCounts),
            nonBcwdbFillByDay: isPartiallyAllocated
              ? p.days.map((_, i) => Math.max(0, 8 - berksHoursOnDay(p, bcwdbFunds, nonBcwdbFunds, i)))
              : null,
          };
        })
        .sort((a, b) => a.start - b.start),
    };
  });

  groupedEmployees.sort((a, b) => a.employeeName.localeCompare(b.employeeName));

  return { groupedEmployees, warnings, periods };
}

/** Total Berks hours on day `i` of a period: holiday + PTO + every fund. */
function berksHoursOnDay(period, bcwdbFunds, nonBcwdbFunds, i) {
  return (
    period.holidayByDay[i] +
    period.ptoByDay[i] +
    bcwdbFunds.reduce((s, f) => s + f.byDay[i], 0) +
    nonBcwdbFunds.reduce((s, f) => s + f.byDay[i], 0)
  );
}

/** Returns the key with the highest count in a Map(name -> count), or null if the map is empty. */
function mostFrequent(counts) {
  let best = null;
  let bestCount = 0;
  for (const [name, count] of counts) {
    if (count > bestCount) {
      best = name;
      bestCount = count;
    }
  }
  return best;
}
