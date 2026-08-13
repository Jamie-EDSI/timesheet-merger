// ---------------------------------------------------------------------------
// PROJECT CODE  ->  TIMESHEET FUND MAPPING
// ---------------------------------------------------------------------------
// This is the one place to edit if a new Project Code shows up in a check
// report, or if a fund name needs to change. Each entry needs:
//
//   label      - the exact text that will be written into the timesheet's
//                fund column (matches the template's dropdown list when
//                possible - see "matchesDropdown" below).
//   category   - "bcwdb"   -> goes in the HOURS WORKED BCWDB columns (G-J)
//                "nonBcwdb"-> goes in the HOURS WORKED NON-BCWDB columns (K-M)
//                "leave"   -> goes in the LEAVE TIME columns (Holiday/PTO)
//                "unmapped"-> doesn't match anything on the timesheet; the
//                             app will still write the hours in so nothing
//                             gets lost, but it will flag it as a warning.
//   leaveColumn - only used when category === "leave": "holiday" or "pto"
//   matchesDropdown - true if `label` is an exact value in the template's
//                dropdown list, false if we normalized/guessed at it.
//
// The timesheet's own dropdown lists (from Contractor_Timesheet_PY26.xlsx):
//   BCWDB:      WIOA A, WIOA DW, WIOA Y, TANF Y, WIOA PWE, BST, EARN, SNAP,
//               OSO, ComOut, VEAP, Business Services
//   Non-BCWDB:  UWBC RSD, UWBC GM, RSD
//
// "Business Services" was added to the BCWDB dropdown (Q24 on Sheet1,
// extending the G11:J11 list's source range from $Q$13:$Q$23 to
// $Q$13:$Q$24 - see public/template/Contractor_Timesheet_PY26.xlsx) so that
// Project Code 525001 could stop being flagged as unmapped; Morgan confirmed
// staff legitimately charge time to it.
// ---------------------------------------------------------------------------

export const PROJECT_CODE_MAP = {
  "525001": {
    label: "Business Services",
    category: "bcwdb",
    matchesDropdown: true,
  },
  "525002": {
    label: "EARN",
    category: "bcwdb",
    matchesDropdown: true,
  },
  "525006": {
    label: "WIOA A",
    category: "bcwdb",
    matchesDropdown: true,
    note: 'Normalized from "WIOA Adult" to match the template\'s dropdown value "WIOA A".',
  },
  "525007": {
    label: "WIOA DW",
    category: "bcwdb",
    matchesDropdown: true,
  },
  "525008": {
    label: "OSO",
    category: "bcwdb",
    matchesDropdown: true,
    note: 'Normalized from "One Stop Operator" to match the template\'s dropdown value "OSO".',
  },
  "525010": {
    label: "PTO",
    category: "leave",
    leaveColumn: "pto",
    matchesDropdown: true,
    note: "Routed to the LEAVE TIME > PTO column instead of a fund column.",
  },
  "525011": {
    label: "SNAP",
    category: "bcwdb",
    matchesDropdown: true,
  },
  "525014": {
    label: "ComOut",
    category: "bcwdb",
    matchesDropdown: true,
    note: 'Normalized from "Community Outreach" to match the template\'s dropdown value "ComOut".',
  },
  "525015": {
    label: "VEAP",
    category: "bcwdb",
    matchesDropdown: true,
    note: 'Normalized from "VEP Grant" to match the template\'s dropdown value "VEAP" (closest existing option).',
  },
};

export function lookupProjectCode(code) {
  const entry = PROJECT_CODE_MAP[code];
  if (entry) return entry;
  return {
    label: `Unknown Code ${code}`,
    category: "unmapped",
    matchesDropdown: false,
    note: `Project Code ${code} was not in the mapping table at all (lib/fundMap.js). Hours are written in as plain text - add this code to PROJECT_CODE_MAP to fix it properly.`,
  };
}
