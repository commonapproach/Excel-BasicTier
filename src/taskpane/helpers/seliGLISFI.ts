/* global Excel, console */
/* eslint-disable no-console */
import { fetchAndParseSeliGLI, SeliGLIData } from "../domain/fetchServer/getSeliGLISFI";

interface SeliIndicatorExtended {
  // Allow derived field access without excessive type narrowing
  [key: string]: any; // eslint-disable-line @typescript-eslint/no-explicit-any
  "@id": string;
  hasName: string;
  hasDescription?: string;
  forOutcome: string;
  forTheme?: string; // derived if missing
}

export async function populateSeliGLISFI(): Promise<{
  themes: number;
  outcomes: number;
  indicators: number;
}> {
  const { themes, outcomes, indicators }: SeliGLIData = await fetchAndParseSeliGLI();

  // Outcome -> Theme map to derive Indicator.forTheme
  const outcomeThemeMap: Record<string, string> = {};
  outcomes.forEach((o) => {
    if (o.forTheme) {
      outcomeThemeMap[o["@id"]] = o.forTheme;
    }
  });

  const ESDC_ID = "https://ontology.commonapproach.org/cids#esdc";
  const ESDC_NAME = "Employment and Social Development Canada";

  await Excel.run(async (context) => {
    await upsertEntity(context, "Organization", [{ "@id": ESDC_ID, "hasLegalName": ESDC_NAME }], {
      simple: ["@id", "hasLegalName"],
      links: [],
    });
  });

  const augmentedIndicators: SeliIndicatorExtended[] = indicators.map((ind) => {
    if (ind.forOutcome && outcomeThemeMap[ind.forOutcome] && !(ind as any).forTheme) {
      return { ...ind, forTheme: outcomeThemeMap[ind.forOutcome], forOrganization: ESDC_ID };
    }
    return { ...ind, forOrganization: ESDC_ID } as SeliIndicatorExtended;
  });

  await Excel.run(async (context) => {
    await upsertEntity(context, "Theme", themes, {
      simple: ["@id", "hasName", "hasDescription"],
      links: [],
    });
    await upsertEntity(context, "Outcome", outcomes.map(o => ({ ...o, forOrganization: ESDC_ID })), {
      simple: ["@id", "hasName", "hasDescription"],
      links: ["forTheme", "hasIndicator", "forOrganization"],
    });
    await upsertEntity(context, "Indicator", augmentedIndicators, {
      simple: ["@id", "hasName", "hasDescription"],
      links: ["forOutcome", "forTheme", "forOrganization"],
    });
    try {
      await rebuildOutcomeIndicatorLinks(context);
    } catch (err) {
      console.warn("Failed to rebuild outcome hasIndicator lists:", err);
    }
  });

  return { themes: themes.length, outcomes: outcomes.length, indicators: indicators.length };
}

async function upsertEntity(
  context: Excel.RequestContext,
  tableName: string,
  dataArr: any[], // eslint-disable-line @typescript-eslint/no-explicit-any
  config: { simple: string[]; links: string[] }
) {
  let worksheet: Excel.Worksheet;
  try {
    worksheet = context.workbook.worksheets.getItem(tableName);
  } catch {
    console.warn(`Worksheet ${tableName} not found; skipping.`);
    return;
  }

  let table: Excel.Table;
  try {
    table = worksheet.tables.getItem(tableName);
  } catch {
    console.warn(`Table ${tableName} not found on worksheet; skipping.`);
    return;
  }

  const headerRange = table.getHeaderRowRange();
  headerRange.load("values");
  await context.sync();
  const headers: string[] = headerRange.values[0].map((h: string) => (h === "'@id" ? "@id" : h));

  const colIndex: Record<string, number> = {};
  [...config.simple, ...config.links].forEach((f) => {
    const i = headers.indexOf(f);
    if (i !== -1) colIndex[f] = i;
  });
  const idIdx = headers.indexOf("@id");
  if (idIdx === -1) {
    console.warn(`Table ${tableName} lacks @id column; skipping.`);
    return;
  }

  let dataBody: Excel.Range | null = null;
  try {
    dataBody = table.getDataBodyRange();
    dataBody.load(["values", "rowCount"]);
    await context.sync();
  } catch {
    // empty table
  }
  const values: any[][] = dataBody ? (dataBody.values as any[][]) : []; // eslint-disable-line @typescript-eslint/no-explicit-any
  const rowCount = dataBody ? dataBody.rowCount : 0;

  const idToRow: Record<string, number> = {};
  for (let r = 0; r < rowCount; r++) {
    const idVal = values[r] && values[r][idIdx];
    if (idVal) idToRow[String(idVal)] = r;
  }

  function findEmptyRow(): number {
    if (!values.length) return 0;
    for (let r = 0; r < values.length; r++) {
      if (!values[r] || !values[r][idIdx]) return r;
    }
    return values.length;
  }

  for (const item of dataArr) {
    const itemId = item["@id"] || item.id;
    if (!itemId) continue;
    let targetRow = idToRow[itemId];
    if (targetRow === undefined) {
      targetRow = findEmptyRow();
      if (!values[targetRow]) values[targetRow] = new Array(headers.length).fill("");
      values[targetRow][idIdx] = itemId;
      idToRow[itemId] = targetRow;
    }
    for (const f of config.simple) {
      if (f === "@id") continue;
      if (colIndex[f] !== undefined && item[f] !== undefined) {
        values[targetRow][colIndex[f]] = item[f];
      }
    }
    for (const lf of config.links) {
      if (colIndex[lf] === undefined) continue;
      const val = item[lf];
      if (!val) continue;
      values[targetRow][colIndex[lf]] = Array.isArray(val) ? val.join(", ") : val;
    }
  }

  if (values.length) {
    const writeRange = table.getDataBodyRange();
    writeRange.values = values;
  }
  await context.sync();
}

async function rebuildOutcomeIndicatorLinks(context: Excel.RequestContext) {
  let indTable: Excel.Table;
  let outTable: Excel.Table;
  try {
    indTable = context.workbook.worksheets.getItem("Indicator").tables.getItem("Indicator");
  } catch {
    return;
  }
  try {
    outTable = context.workbook.worksheets.getItem("Outcome").tables.getItem("Outcome");
  } catch {
    return;
  }

  const indHeadersR = indTable.getHeaderRowRange();
  const outHeadersR = outTable.getHeaderRowRange();
  indHeadersR.load("values");
  outHeadersR.load("values");
  await context.sync();

  const indHeaders = indHeadersR.values[0].map((h: string) => (h === "'@id" ? "@id" : h));
  const outHeaders = outHeadersR.values[0].map((h: string) => (h === "'@id" ? "@id" : h));
  const indIdIdx = indHeaders.indexOf("@id");
  const indForOutcomeIdx = indHeaders.indexOf("forOutcome");
  const outIdIdx = outHeaders.indexOf("@id");
  const outHasIndicatorIdx = outHeaders.indexOf("hasIndicator");
  if (indIdIdx === -1 || indForOutcomeIdx === -1 || outIdIdx === -1 || outHasIndicatorIdx === -1) {
    return;
  }

  let indData: Excel.Range;
  let outData: Excel.Range;
  try {
    indData = indTable.getDataBodyRange();
  } catch {
    return;
  }
  try {
    outData = outTable.getDataBodyRange();
  } catch {
    return;
  }
  indData.load(["values", "rowCount"]);
  outData.load(["values", "rowCount"]);
  await context.sync();

  const indValues = indData.values as any[][]; // eslint-disable-line @typescript-eslint/no-explicit-any
  const outValues = outData.values as any[][]; // eslint-disable-line @typescript-eslint/no-explicit-any
  const indRowCount = indData.rowCount;
  const outRowCount = outData.rowCount;

  const outIndicators: Record<string, Set<string>> = {};
  for (let r = 0; r < indRowCount; r++) {
    const row = indValues[r];
    if (!row || !row[indIdIdx]) continue;
    const indId = String(row[indIdIdx]);
    const fo = row[indForOutcomeIdx] ? String(row[indForOutcomeIdx]) : "";
    if (!fo) continue;
    if (!outIndicators[fo]) outIndicators[fo] = new Set();
    outIndicators[fo].add(indId);
  }

  let updates = 0;
  for (let r = 0; r < outRowCount; r++) {
    const row = outValues[r];
    if (!row || !row[outIdIdx]) continue;
    const outId = String(row[outIdIdx]);
    const indicatorsSet = outIndicators[outId];
    const newVal = indicatorsSet ? Array.from(indicatorsSet).join(", ") : "";
    if (row[outHasIndicatorIdx] !== newVal) {
      outData.getCell(r, outHasIndicatorIdx).values = [[newVal]];
      updates++;
    }
  }
  if (updates) await context.sync();
}
