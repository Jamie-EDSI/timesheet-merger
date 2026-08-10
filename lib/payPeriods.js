// ---------------------------------------------------------------------------
// Pay period helpers
// ---------------------------------------------------------------------------
// The timesheet template covers a fixed 2-week block: Mon-Fri, then the
// following Mon-Fri (10 working days). A check report can span multiple
// weeks (the sample Berks report spans 4 weeks / two pay periods).
//
// We don't have an actual payroll calendar to consult, so we anchor to the
// Monday of the EARLIEST date seen anywhere in the report, then bucket every
// date into consecutive, non-overlapping 14-day windows from that anchor.
// This matches the sample data exactly: 3/30 (Mon) - 4/10 (Fri) is period 1,
// 4/13 (Mon) - 4/24 (Fri) is period 2.
//
// If your organization's actual pay period calendar doesn't start on the
// Monday of the first date in a report, this is the file to adjust.
// ---------------------------------------------------------------------------

const MS_PER_DAY = 24 * 60 * 60 * 1000;

/** Returns a new Date set to midnight local time for the given date. */
export function atMidnight(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/** Returns the Monday (as a Date) of the week containing `date`. */
export function mondayOf(date) {
  const d = atMidnight(date);
  const day = d.getDay(); // 0 = Sun, 1 = Mon, ... 6 = Sat
  const diff = day === 0 ? -6 : 1 - day; // shift back to Monday
  return addDays(d, diff);
}

export function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}

export function diffDays(a, b) {
  return Math.round((atMidnight(a) - atMidnight(b)) / MS_PER_DAY);
}

export function dateKey(date) {
  const d = atMidnight(date);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(
    d.getDate()
  ).padStart(2, "0")}`;
}

export function formatMDY(date) {
  const d = atMidnight(date);
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}

/**
 * Given an array of JS Dates, returns:
 *   {
 *     anchor: Date,                     // Monday of the earliest date
 *     periodOf: (date) => periodKey,    // dateKey of that period's Monday
 *     periods: [{ key, start, end, days: [10 Dates] }, ...]  // sorted
 *   }
 */
export function computePayPeriods(dates) {
  if (!dates.length) {
    return { anchor: null, periodOf: () => null, periods: [] };
  }
  const earliest = dates.reduce((min, d) => (d < min ? d : min), dates[0]);
  const anchor = mondayOf(earliest);

  const periodStartFor = (date) => {
    const offsetDays = diffDays(mondayOf(date), anchor);
    const periodIndex = Math.floor(offsetDays / 14);
    return addDays(anchor, periodIndex * 14);
  };

  const periodsByKey = new Map();
  for (const date of dates) {
    const start = periodStartFor(date);
    const key = dateKey(start);
    if (!periodsByKey.has(key)) {
      const end = addDays(start, 11); // 2nd Friday
      periodsByKey.set(key, { key, start, end, days: buildTenDays(start) });
    }
  }

  const periods = Array.from(periodsByKey.values()).sort((a, b) => a.start - b.start);

  return {
    anchor,
    periodOf: (date) => dateKey(periodStartFor(date)),
    periods,
  };
}

/** Builds the 10 working days (Mon-Fri, Mon-Fri) for a period starting on `start` (a Monday). */
function buildTenDays(start) {
  const days = [];
  for (let week = 0; week < 2; week++) {
    for (let d = 0; d < 5; d++) {
      days.push(addDays(start, week * 7 + d));
    }
  }
  return days;
}

/** Converts a JS Date to an Excel date serial number (Excel's 1900 date system, with the leap-year bug baked in like Excel itself). */
export function toExcelSerial(date) {
  const d = atMidnight(date);
  const excelEpoch = new Date(1899, 11, 30); // Excel day 0
  return diffDays(d, excelEpoch);
}
