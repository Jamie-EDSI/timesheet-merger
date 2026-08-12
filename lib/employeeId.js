// ---------------------------------------------------------------------------
// Employee ID normalization
// ---------------------------------------------------------------------------
// Newer check reports zero-pad Employee IDs (e.g. "0000971"), while older
// reports (and older tabs already written to an "ongoing" workbook) use the
// bare number ("971"). Stripping leading zeros everywhere an Employee ID is
// read - whether from a freshly parsed report row or from a marker cell on
// an existing tab - keeps the two formats matching instead of producing a
// duplicate tab for the same person.
// ---------------------------------------------------------------------------

/** Strips leading zeros from an Employee ID. Never returns "" for a non-empty input (e.g. "0000" -> "0"). */
export function normalizeEmployeeId(id) {
  const trimmed = String(id ?? "").trim();
  if (!trimmed) return trimmed;
  const stripped = trimmed.replace(/^0+/, "");
  return stripped === "" ? "0" : stripped;
}
