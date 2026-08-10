import { computePayPeriods, dateKey } from "./payPeriods.js";
import { lookupProjectCode } from "./fundMap.js";

/**
 * Takes the flat list of parsed report rows and turns it into the structure
 * lib/timesheetBuilder.js needs: one entry per employee, each with a list
 * of pay periods, each period holding per-day hours already sorted into
 * "holiday", "pto", "bcwdb funds", and "non-bcwdb funds".
 */
export function groupEntries(entries) {
  const warnings = [];
  const unmappedCodesSeen = new Map(); // code -> {label, note}

  const allDates = entries.map((e) => e.date);
  const { periods, periodOf } = computePayPeriods(allDates);
  const periodByKey = new Map(periods.map((p) => [dateKey(p.start), p]));

  // employeeId -> { name, periods: Map(periodKey -> workingPeriod) }
  const employees = new Map();

  for (const entry of entries) {
    const fund = lookupProjectCode(entry.projectCode);
    if (fund.category === "unmapped" && !unmappedCodesSeen.has(entry.projectCode)) {
      unmappedCodesSeen.set(entry.projectCode, fund);
    }

    if (!employees.has(entry.employeeId)) {
      employees.set(entry.employeeId, { name: entry.employeeName, periods: new Map() });
    }
    const employee = employees.get(entry.employeeId);
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
      });
    }
    const period = employee.periods.get(periodKey);
    const dayIndex = period.days.findIndex((d) => dateKey(d) === dateKey(entry.date));
    if (dayIndex === -1) {
      warnings.push(
        `${employee.name} (ID ${entry.employeeId}): a row dated ${entry.date.toDateString()} didn't line up with a weekday in its pay period (maybe a weekend date in the report?) and was skipped.`
      );
      continue;
    }

    if (fund.category === "leave") {
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

  const groupedEmployees = Array.from(employees.entries()).map(([employeeId, employee]) => ({
    employeeId,
    employeeName: employee.name,
    periods: Array.from(employee.periods.values())
      .map((p) => ({
        ...p,
        bcwdbFunds: Array.from(p.bcwdbFundsByLabel.values()),
        nonBcwdbFunds: Array.from(p.nonBcwdbFundsByLabel.values()),
      }))
      .sort((a, b) => a.start - b.start),
  }));

  groupedEmployees.sort((a, b) => a.employeeName.localeCompare(b.employeeName));

  return { groupedEmployees, warnings, periods };
}
