import moment from "moment-timezone";

/**
 * Pseudo-year bucket for indicator reports that have no parseable
 * prov:endedAtTime / prov:startedAtTime value.
 */
export const UNDATED_YEAR = "Undated";

export interface OrgExportInfo {
  /** Organization @id (the value stored in IndicatorReport.forOrganization). */
  id: string;
  /** Display name (org:hasLegalName), falling back to the @id for orphans. */
  name: string;
  /** Year (string) -> number of indicator reports for this org in that year. */
  reportsByYear: Record<string, number>;
}

/**
 * Converts an Excel cell value (date serial number or string) to a 4-digit
 * year string. Returns null when the value is empty or not a valid date.
 *
 * Uses the same serial-number conversion and local-timezone handling as the
 * export pipeline so the bucketed year matches the exported prov: timestamps.
 */
export function excelDateToYear(value: any): string | null {
  if (value === null || value === undefined || value === "") return null;

  let dateInput: any = value;
  if (typeof value === "number") {
    dateInput = new Date((value - (25567 + 1)) * 86400 * 1000);
  }

  const localTimezone = moment.tz.guess();
  const parsed = moment(dateInput).tz(localTimezone);
  if (!parsed.isValid()) return null;

  return String(parsed.year());
}

/* global Excel */
/**
 * Scans the workbook for organizations that have at least one indicator
 * report. Returns one entry per such organization, with a per-year report
 * count used to drive the export wizard's organization and year steps.
 */
export async function getExportableOrganizations(): Promise<OrgExportInfo[]> {
  return await Excel.run(async (context: Excel.RequestContext) => {
    const workbook = context.workbook;
    workbook.load("tables");
    await context.sync();

    const tables = workbook.tables.items;
    const orgTable = tables.find((t) => t.name === "Organization");
    const reportTable = tables.find((t) => t.name === "IndicatorReport");

    // Without these two tables there is nothing to scope an export to.
    if (!orgTable || !reportTable) return [];

    const orgHeader = orgTable.getHeaderRowRange();
    orgHeader.load("values");
    const orgBody = orgTable.getDataBodyRange();
    orgBody?.load("values");

    const repHeader = reportTable.getHeaderRowRange();
    repHeader.load("values");
    const repBody = reportTable.getDataBodyRange();
    repBody?.load("values");

    await context.sync();

    // Build org @id -> legal name lookup.
    const orgNames: Record<string, string> = {};
    const orgHeaders = (orgHeader.values[0] || []) as string[];
    const orgIdIdx = orgHeaders.indexOf("@id");
    const orgNameIdx = orgHeaders.indexOf("hasLegalName");
    if (orgBody && orgIdIdx !== -1) {
      for (const row of orgBody.values) {
        const id = String(row[orgIdIdx] ?? "").trim();
        if (!id) continue;
        const name = orgNameIdx !== -1 ? String(row[orgNameIdx] ?? "").trim() : "";
        orgNames[id] = name;
      }
    }

    // Tally indicator reports per organization per year.
    const result: Record<string, OrgExportInfo> = {};
    const repHeaders = (repHeader.values[0] || []) as string[];
    const forOrgIdx = repHeaders.indexOf("forOrganization");
    const endedIdx = repHeaders.indexOf("endedAtTime");
    const startedIdx = repHeaders.indexOf("startedAtTime");

    if (repBody && forOrgIdx !== -1) {
      for (const row of repBody.values) {
        const orgId = String(row[forOrgIdx] ?? "").trim();
        if (!orgId) continue;

        const year =
          excelDateToYear(endedIdx !== -1 ? row[endedIdx] : null) ??
          excelDateToYear(startedIdx !== -1 ? row[startedIdx] : null) ??
          UNDATED_YEAR;

        if (!result[orgId]) {
          result[orgId] = {
            id: orgId,
            name: orgNames[orgId] || orgId,
            reportsByYear: {},
          };
        }
        result[orgId].reportsByYear[year] = (result[orgId].reportsByYear[year] ?? 0) + 1;
      }
    }

    return Object.values(result).sort((a, b) => a.name.localeCompare(b.name));
  });
}
