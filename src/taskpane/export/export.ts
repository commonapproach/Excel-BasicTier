import moment from "moment-timezone";
import { IntlShape } from "react-intl";
import { CodeList, getCodeListByTableName } from "../domain/fetchServer/getCodeLists";
import { UNIT_DEFINITIONS, UNIT_IRI, getUnitDefinition } from "../domain/fetchServer/getUnitsOfMeasure";
import { TableInterface } from "../domain/interfaces/table.interface";
import {
  contextUrl,
  createInstance,
  ignoredFields,
  map,
  mapSFFModel,
  ModelType,
  predefinedCodeLists,
  SFFModelType,
} from "../domain/models";
import { Base as BaseModel, FieldType } from "../domain/models/Base";
import { validate } from "../domain/validation/validator";
import { downloadJSONLD } from "../utils/utils";

/* global Excel*/
export async function exportData(
  intl: IntlShape,
  orgName: string,
  setDialogContent: (header: string, content: string, nextCallBack?: Function) => void
): Promise<void> {
  await Excel.run(async (context: Excel.RequestContext) => {
    // Get the tables from the workbook
    const workbook = context.workbook;
    workbook.load("tables");
    await context.sync();
    const tables = workbook.tables.items;
    const data: TableInterface[] = [];
    const changeOnDefaultCodeListsWarning: string[] = []; // Array to store warnings about code list changes

    let fullMap = map;

    // Check if any table of the SFF module is created
    const tableNamesOnBase = tables.map((table) => table.name);
    const sffModuleTables = Object.keys(mapSFFModel);
    if (sffModuleTables.some((table) => tableNamesOnBase.includes(table))) {
      fullMap = { ...map, ...mapSFFModel };
    }

    const tableNames = tables.map((item) => item.name);
    for (const [key] of Object.entries(fullMap)) {
      if (!tableNames.includes(key)) {
        setDialogContent(
          intl.formatMessage({
            id: "generics.error",
            defaultMessage: "Error",
          }),
          intl.formatMessage(
            {
              id: "export.messages.error.missingTable",
              defaultMessage:
                "Table <b>{tableName}</b> is missing. Please create the tables first.",
            },
            { tableName: key, b: (str) => `<b>${str}</b>` }
          )
        );
        return;
      }
    }

    // Pre-fetch all code lists we'll need in one batch
    const codeListCache: Record<string, CodeList[]> = {};
    const codeListPromises = predefinedCodeLists
      .filter(tableNames.includes.bind(tableNames))
      .map(async (tableName) => {
        codeListCache[tableName] = await getCodeListByTableName(tableName);
      });

    await Promise.all(codeListPromises);

    // First, prepare all the table objects we'll work with and queue all load operations
    const tablesToProcess = tables.filter((table) => Object.keys(fullMap).includes(table.name));
    const tableInfos = tablesToProcess.map((table) => {
      const tableRange = table.getRange();
      const tableHeaderRange = tableRange.getRow(0);

      // Load all the properties we'll need
      tableHeaderRange.load("values");
      table.load("name");

      // For data, we need to load the data body range separately
      const dataBodyRange = table.getDataBodyRange();
      dataBodyRange?.load("values");

      return {
        table,
        tableHeaderRange,
        dataBodyRange,
        codeList: codeListCache[table.name],
      };
    });

    // Execute a single sync to load all data
    await context.sync();

    // Now process each table with loaded data
    for (const { table, tableHeaderRange, dataBodyRange, codeList } of tableInfos) {
      if (!dataBodyRange) continue; // Skip if no data (empty table)

      const headers = tableHeaderRange.values[0];
      const records = dataBodyRange.values;
      const idColumnIndex = headers.indexOf("@id");
      const cid: BaseModel = createInstance(table.name as ModelType | SFFModelType);
      const tableFields = cid.getTopLevelFields();

      // Process all records for this table
      for (const recordValues of records) {
        // Skip records that are defined in the common approach code lists
        if (
          codeList &&
          idColumnIndex !== -1 &&
          codeList.find((item) => item["@id"] === recordValues[idColumnIndex])
        ) {
          // Check if the record has changes compared to the code list item
          const recordId = recordValues[idColumnIndex];
          const existingItem = codeList.find((item) => item["@id"] === recordId);

          if (existingItem) {
            let hasChanges = false;
            for (const fieldName of Object.keys(existingItem)) {
              const fieldIndex = headers.indexOf(fieldName);
              if (fieldIndex !== -1) {
                const recordValue = recordValues[fieldIndex];
                const existingValue = (existingItem as Record<string, any>)[fieldName];

                // Compare values
                if (
                  recordValue !== undefined &&
                  recordValue !== null &&
                  recordValue.toString() !== existingValue?.toString()
                ) {
                  hasChanges = true;
                  break;
                }
              }
            }

            if (hasChanges) {
              changeOnDefaultCodeListsWarning.push(
                intl.formatMessage(
                  {
                    id: "export.messages.warning.codeListChangesIgnored",
                    defaultMessage:
                      "Changes made in the predefined code list item with @id <b>{id}</b> in table <b>{tableName}</b> will be ignored.",
                  },
                  {
                    id: recordId,
                    tableName: table.name,
                    b: (str: string) => `<b style="word-break: break-word;">${str}</b>`,
                  }
                ) as string
              );
            }
          }
          continue;
        }

        // Check if record is similar to code list item (all values match except @id)
        if (codeList) {
          // Find a matching code list item where all fields except @id match
          const similarItem = codeList.find((item) =>
            Object.keys(item).every((key) => {
              if (key === "@id") return true;

              const fieldIndex = headers.indexOf(key);
              if (fieldIndex === -1) return true;

              const recordValue = recordValues[fieldIndex];
              const existingValue = (item as Record<string, any>)[key];

              return recordValue?.toString() === existingValue?.toString();
            })
          );

          if (similarItem) {
            const recordIdIndex = headers.indexOf("@id");
            const recordId = recordIdIndex !== -1 ? recordValues[recordIdIndex] : "";

            changeOnDefaultCodeListsWarning.push(
              intl.formatMessage(
                {
                  id: "export.messages.warning.codeListSimilarItem",
                  defaultMessage:
                    "Record in table <b>{tableName}</b> with @id: <b>{recordId}</b> is similar to the predefined code list item with @id: <b>{codeListItemId}</b>.<br/>Please review the code list item before exporting, or a custom code list item will be exported.",
                },
                {
                  codeListItemId: similarItem["@id"],
                  recordId: recordId || "",
                  tableName: table.name,
                  b: (str: string) => `<b style="word-break: break-word;">${str}</b>`,
                }
              ) as string
            );
          }
        }

        // Row initialization advanced logic to mirror latest Airtable extension:
        // Population uses i72:Population; SFF module tables use sff: prefix; Address uses ic:; others cids:
        const isSFFTable = Object.prototype.hasOwnProperty.call(mapSFFModel, table.name);
        const computedType =
          table.name === "Population"
            ? "i72:Population"
            : isSFFTable
            ? `sff:${table.name}`
            : `cids:${table.name}`;
        const row: TableInterface = {
          "@context": contextUrl,
          "@type": computedType,
          "@id": "",
        };

        const isEmpty = true; // Flag to check if the row is empty

        // Process all fields in one pass using the cached headers
        await processRecord(tableFields, headers, recordValues, row, isEmpty).then((result) => {
          const [processedRow, rowIsEmpty] = result;
          if (!rowIsEmpty) {
            data.push(processedRow);
          }
        });
      }
    }

    // Post-processing enhancements (multi-typing, units, cleaning) before validation & warnings

    // Add multi-typing for Indicators that have i72:cardinality_of link (become i72:Cardinality as additional @type)
    for (const item of data) {
      const typeVal = item["@type"];
      if (!typeVal) continue;
      const types = Array.isArray(typeVal) ? [...typeVal] : [typeVal];
      const isIndicator = types.includes("cids:Indicator");
      if (isIndicator && item["i72:cardinality_of"]) {
        if (!types.includes("i72:Cardinality")) {
          item["@type"] = [...types, "i72:Cardinality"]; // mutate
        }
      }
    }

    // Ensure each Indicator has a unit_of_measure (default unspecified) & propagate to IndicatorReport value objects
    const indicatorUnitById: Record<string, string> = {};
    const usedUnitIris: Set<string> = new Set();
    for (const item of data) {
      const typeVal = item["@type"]; const types = Array.isArray(typeVal) ? typeVal : [typeVal];
      if (types.includes("cids:Indicator")) {
        if (item["@id"]) {
          const existing = item["i72:unit_of_measure"] as string | undefined;
          const resolved = existing && existing.trim() !== "" ? existing : UNIT_IRI.UNSPECIFIED;
            if (!existing) item["i72:unit_of_measure"] = resolved;
          indicatorUnitById[item["@id"] as string] = resolved;
          usedUnitIris.add(resolved);
        }
      }
    }
    for (const item of data) {
      const typeVal = item["@type"]; const types = Array.isArray(typeVal) ? typeVal : [typeVal];
      if (types.includes("cids:IndicatorReport")) {
        const indicatorId = item["forIndicator"]; // link field name assumption
        const valueObj = item["i72:value"] as Record<string, any> | undefined;
        if (valueObj && !valueObj["i72:unit_of_measure"]) {
          const fallback = (typeof indicatorId === "string" && indicatorUnitById[indicatorId]) || UNIT_IRI.UNSPECIFIED;
          valueObj["i72:unit_of_measure"] = fallback;
          usedUnitIris.add(fallback);
        }
      }
    }

    // Inject unit definition objects for any used unit IRIs (avoid duplicates) including related cids unit IRIs referenced
    const queue: string[] = Array.from(usedUnitIris);
    const seen: Set<string> = new Set();
    while (queue.length > 0) {
      const iri = queue.shift() as string;
      if (seen.has(iri)) continue;
      seen.add(iri);
      // Attempt fetch definition (may fall back to static)
      try {
        const def = (await getUnitDefinition(iri)) || UNIT_DEFINITIONS[iri];
        if (def) {
          const already = data.some((d) => d && d["@id"] === iri);
            if (!already) data.push({ "@context": contextUrl, ...def });
          // enqueue nested cids IRIs
          for (const val of Object.values(def)) {
            if (typeof val === "string" && val.startsWith("https://ontology.commonapproach.org/cids#")) {
              queue.push(val);
            }
          }
        }
      } catch { /* ignore */ }
    }

    const { errors, warnings } = await validate(data, "export", intl);

    // Load all data for validation checks at once
    const warningCheckData = await loadDataForWarnings(intl, context, tables, fullMap);
    const noExportingFields = warningCheckData.notExportedFields;
    const emptyTableWarning = warningCheckData.emptyTableWarnings;

    // Include the code list warnings in the warnings
    const allWarnings = [
      ...noExportingFields,
      ...warnings,
      ...emptyTableWarning,
      ...changeOnDefaultCodeListsWarning,
    ].join("<hr/>");

    if (errors.length > 0) {
      setDialogContent(
        intl.formatMessage({
          id: "generics.error",
          defaultMessage: "Error",
        }),
        errors.map((item) => `<p>${item}</p>`).join("")
      );
      return;
    }

  // Always deep-clean just before potentially showing warnings (but keep original for reference)
  const cleanedData = deepCleanExportObjects(data);

  if (allWarnings.length > 0) {
      setDialogContent(
        intl.formatMessage({
          id: "generics.warning",
          defaultMessage: "Warning",
        }),
        allWarnings,
        () => {
          setDialogContent(
            intl.formatMessage({
              id: "generics.warning",
              defaultMessage: "Warning",
            }),
            intl.formatMessage({
              id: "export.messages.warning.continue",
              defaultMessage: "<p>Do you want to export anyway?</p>",
            }),
            () => {
        downloadJSONLD(cleanedData, `${getFileName(orgName)}.json`);
              setDialogContent(
                intl.formatMessage({
                  id: "generics.success",
                  defaultMessage: "Success",
                }),
                intl.formatMessage({
                  id: "export.messages.success",
                  defaultMessage: "Data exported successfully!",
                })
              );
            }
          );
        }
      );
      return;
    }
  downloadJSONLD(cleanedData, `${getFileName(orgName)}.json`);
    setDialogContent(
      intl.formatMessage({
        id: "generics.success",
        defaultMessage: "Success",
      }),
      intl.formatMessage({
        id: "export.messages.success",
        defaultMessage: "Data exported successfully!",
      })
    );
  });
}

// Load all data needed for warning checks in a single batch
async function loadDataForWarnings(
  intl: IntlShape,
  context: Excel.RequestContext,
  tables: Excel.Table[],
  fullMap: any
) {
  const relevantTables = tables.filter((table) => Object.keys(fullMap).includes(table.name));

  // Queue loads for all header ranges and data ranges
  const tableData = relevantTables.map((table) => {
    const headerRange = table.getHeaderRowRange();
    headerRange.load("values");

    const dataRange = table.getDataBodyRange();
    // Only load if the range exists (not an empty table)
    if (dataRange) {
      dataRange.load("values");
    }

    const cid = createInstance(table.name as ModelType | SFFModelType);
    const internalFields = cid.getAllFields().map((item) => item.displayName || item.name);

    return {
      table,
      headerRange,
      dataRange,
      internalFields,
    };
  });

  // Execute a single sync for all loads
  await context.sync();

  // Process the loaded data
  const notExportedFields: string[] = [];
  const emptyTableWarnings: string[] = [];

  for (const { table, headerRange, dataRange, internalFields } of tableData) {
    // Process header information for not exported fields
    const headers = headerRange.values[0];
    for (const field of headers) {
      if (
        Object.keys(fullMap).includes(field) ||
        (ignoredFields as any)[table.name]?.includes(field)
      ) {
        continue;
      }
      if (!internalFields.includes(field)) {
        notExportedFields.push(
          intl.formatMessage(
            {
              id: Object.keys(map).includes(table.name)
                ? "export.messages.warning.fieldWillNotBeExported"
                : "export.messages.warning.notExported",
              defaultMessage:
                "Field <b>{fieldName}</b> on table <b>{tableName}</b> will not be exported",
            },
            { fieldName: field, tableName: table.name, b: (str: string) => `<b>${str}</b>` }
          ) as string
        );
      }
    }

    // Process data range for empty tables
    if (dataRange) {
      const values = dataRange.values;
      let isEmpty = true;

      for (const row of values) {
        if (row.some((cell) => cell)) {
          isEmpty = false;
          break;
        }
      }

      if (isEmpty) {
        emptyTableWarnings.push(
          intl.formatMessage(
            {
              id: "export.messages.warning.emptyTable",
              defaultMessage: "<Table <b>${tableName}</b> is empty",
            },
            {
              tableName: table.name,
              b: (str) => `<b>${str}</b>`,
            }
          )
        );
      }
    } else {
      // If dataRange doesn't exist, the table is empty
      emptyTableWarnings.push(
        intl.formatMessage(
          {
            id: "export.messages.warning.emptyTable",
            defaultMessage: "<Table <b>${tableName}</b> is empty",
          },
          {
            tableName: table.name,
            b: (str) => `<b>${str}</b>`,
          }
        )
      );
    }
  }

  return {
    notExportedFields,
    emptyTableWarnings,
  };
}

// New helper function to process records more efficiently
/* eslint-disable no-param-reassign */
async function processRecord(
  fields: FieldType[],
  headers: string[],
  recordValues: any[],
  row: TableInterface,
  isEmpty: boolean
): Promise<[TableInterface, boolean]> {
  for (const field of fields) {
    const columnIndex = headers.indexOf(field.displayName || field.name);
    const value = recordValues[columnIndex];

    if (field.type === "link") {
      if (field.representedType === "array") {
        const fieldValue = value ?? field?.defaultValue;
        if (fieldValue && fieldValue.length > 0) {
          isEmpty = false;
        }
        row[field.name] =
          typeof fieldValue === "string"
            ? fieldValue.split(", ").filter((v) => v !== "" && v !== null && v !== undefined)
            : (fieldValue as string[]).filter((v) => v !== "" && v !== null && v !== undefined);
      } else if (field.representedType === "string") {
        const fieldValue = value ?? field?.defaultValue;
        if (fieldValue) {
          isEmpty = false;
        }
        row[field.name] = Array.isArray(fieldValue) ? fieldValue[0] : fieldValue;
      }
    } else if (field.type === "object") {
      const [newRow, newIsEmpty] = await getObjectFieldsRecursively(
        headers,
        recordValues,
        field,
        row,
        isEmpty
      );
      row = { ...row, ...newRow };
      isEmpty = newIsEmpty;
    } else if (field.type === "select") {
      const fieldValue = value ?? "";
      if (fieldValue) {
        isEmpty = false;
      }
      let optionField;
      if (field.getOptionsAsync) {
        const options = await field.getOptionsAsync();
        optionField = options.find((opt) => opt.name === fieldValue);
      } else {
        optionField = field.selectOptions?.find((opt) => opt.name === fieldValue);
      }
      if (optionField) {
        row[field.name] = field.representedType === "array" ? [optionField.id] : optionField.id;
      } else {
        row[field.name] = field.defaultValue;
      }
    } else if (field.type === "multiselect") {
      const fieldValue = value ?? "";
      const valuesArray =
        typeof fieldValue === "string"
          ? fieldValue.split(", ").filter((v) => v !== "" && v !== null && v !== undefined)
          : [];
      if (valuesArray.length > 0) {
        isEmpty = false;
      }
      let optionFields = [];
      if (field.getOptionsAsync) {
        const options = await field.getOptionsAsync();
        optionFields = options.filter((opt) => valuesArray.includes(opt.name));
      } else {
        optionFields = field.selectOptions?.filter((opt) => valuesArray.includes(opt.name)) || [];
      }
      if (optionFields.length > 0) {
        row[field.name] =
          field.representedType === "array"
            ? optionFields.map((opt) => opt.id)
            : optionFields.map((opt) => opt.id);
      } else {
        row[field.name] = field.defaultValue;
      }
    } else if (field.type === "datetime") {
      let fieldValue = value ?? "";
      if (fieldValue && (typeof fieldValue === "string" || typeof fieldValue === "number")) {
        isEmpty = false;

        if (typeof fieldValue === "number") {
          // convert excel int date to date
          fieldValue = new Date((fieldValue - (25567 + 1)) * 86400 * 1000);
        }

        // get local timezone
        const localTimezone = moment.tz.guess();
        const date = moment(fieldValue).tz(localTimezone).format("YYYY-MM-DDTHH:mm:ssZ");

        row[field.name] = date;
      } else {
        row[field.name] = "";
      }
    } else if (field.type === "date") {
      let fieldValue = value ?? "";
      if (fieldValue && (typeof fieldValue === "string" || typeof fieldValue === "number")) {
        isEmpty = false;

        if (typeof fieldValue === "number") {
          // convert excel int date to date
          fieldValue = new Date((fieldValue - (25567 + 1)) * 86400 * 1000);
        }

        // get local timezone
        const localTimezone = moment.tz.guess();
        const date = moment(fieldValue).tz(localTimezone).format("YYYY-MM-DD");

        row[field.name] = date;
      } else {
        row[field.name] = "";
      }
    } else if (field.type === "boolean") {
      let fieldValue = value ?? false;

      // Handle string values like "TRUE", "YES", etc.
      if (typeof fieldValue === "string") {
        const upperCaseValue = fieldValue.toUpperCase();
        fieldValue = upperCaseValue === "TRUE" || upperCaseValue === "YES";
      }

      row[field.name] = fieldValue ? true : false;
    } else if (field.type === "number") {
      const fieldValue = value;
      let exportValue: number | null = null;
      if (fieldValue !== null && fieldValue !== undefined && fieldValue !== "") {
        const parsed = Number(fieldValue);
        if (!isNaN(parsed)) {
          exportValue = parsed;
          isEmpty = false;
        }
      }
      row[field.name] = exportValue as any; // number or null handled
    } else {
      const fieldValue = value ?? "";
      if (fieldValue || fieldValue === 0) {
        isEmpty = false;
      }
      let exportValue = fieldValue;
      if (Array.isArray(fieldValue) && field.representedType === "array") {
        exportValue = fieldValue;
      } else if (!Array.isArray(fieldValue) && field.representedType === "array") {
        exportValue = fieldValue ? [fieldValue] : field.defaultValue;
      } else {
        exportValue = fieldValue.toString() || field.defaultValue;
      }
      row[field.name] = exportValue as any; // number or null handled
    }
  }

  return [row, isEmpty];
}
/* eslint-enable no-param-reassign */

function getFileName(orgName: string): string {
  const date = new Date();

  // Get the year, month, and day from the date
  const year = date.getFullYear();
  const month = date.getMonth() + 1; // Add 1 because months are 0-indexed.
  const day = date.getDate();

  // Format month and day to ensure they are two digits
  const monthFormatted = month < 10 ? "0" + month : month;
  const dayFormatted = day < 10 ? "0" + day : day;

  // Concatenate the components to form the desired format (YYYYMMDD)
  const timestamp = `${year}${monthFormatted}${dayFormatted}`;

  return `CIDSBasic${orgName}${timestamp}`;
}

// Deep clean export objects removing null/undefined/empty strings/empty arrays or objects.
// Preserve empty string for i72:hasNumericalValue.
function deepCleanExportObjects(items: TableInterface[]): TableInterface[] {
  const keepEmptyKey = (key: string) => key === "i72:hasNumericalValue";
  const clean = (value: any, parentKey?: string): any => {
    if (Array.isArray(value)) {
      const arr = value.map((v) => clean(v)).filter((v) => {
        if (v === null || v === undefined) return false;
        if (Array.isArray(v) && v.length === 0) return false;
        if (typeof v === "object" && !Array.isArray(v) && Object.keys(v).length === 0) return false;
        return true;
      });
      return arr;
    }
    if (value && typeof value === "object") {
      const objEntries = Object.entries(value)
        .map(([k, v]) => [k, clean(v, k)] as [string, any])
        .filter(([k, v]) => {
          if (v === null || v === undefined) return false;
          if (typeof v === "string" && v.trim() === "" && !keepEmptyKey(k)) return false;
          if (Array.isArray(v) && v.length === 0) return false;
          if (typeof v === "object" && !Array.isArray(v) && Object.keys(v).length === 0) return false;
          return true;
        });
      return Object.fromEntries(objEntries);
    }
    if (typeof value === "string" && value.trim() === "" && !keepEmptyKey(parentKey || "")) {
      return undefined;
    }
    return value;
  };
  return items
    .map((item) => clean(item))
    .filter((i) => i && typeof i === "object" && Object.keys(i).length > 0) as TableInterface[];
}

/* eslint-disable no-param-reassign */
async function getObjectFieldsRecursively(
  headers: string[],
  values: any[],
  field: FieldType,
  row: any,
  isEmpty: boolean
) {
  if (field.type !== "object") {
    const columnIndex = headers.indexOf(field.displayName || field.name);
    const fieldValueOnTable = values[columnIndex];
    if (field.type === "link") {
      if (field.representedType === "array") {
        const fieldValue = fieldValueOnTable ?? field?.defaultValue;
        if (fieldValue && fieldValue.length > 0) {
          isEmpty = false;
        }
        row[field.name] = fieldValue;
      } else if (field.representedType === "string") {
        const fieldValue = fieldValueOnTable ? fieldValueOnTable[0]?.name : field?.defaultValue;
        if (fieldValue) {
          isEmpty = false;
        }
        row[field.name] = fieldValue.toString();
      }
    } else if (field.type === "datetime") {
      let fieldValue = fieldValueOnTable ?? "";
      if (fieldValue && (typeof fieldValue === "string" || typeof fieldValue === "number")) {
        isEmpty = false;

        if (typeof fieldValue === "number") {
          // convert excel int date to date
          fieldValue = new Date((fieldValue - (25567 + 1)) * 86400 * 1000);
        }

        // get local timezone
        const localTimezone = moment.tz.guess();
        const date = moment(fieldValue).tz(localTimezone).format("YYYY-MM-DDTHH:mm:ssZ");

        row[field.name] = date;
      } else {
        row[field.name] = "";
      }
    } else if (field.type === "date") {
      let fieldValue = fieldValueOnTable ?? "";
      if (fieldValue && (typeof fieldValue === "string" || typeof fieldValue === "number")) {
        isEmpty = false;

        if (typeof fieldValue === "number") {
          // convert excel int date to date
          fieldValue = new Date((fieldValue - (25567 + 1)) * 86400 * 1000);
        }

        // get local timezone
        const localTimezone = moment.tz.guess();
        const date = moment(fieldValue).tz(localTimezone).format("YYYY-MM-DD");

        row[field.name] = date;
      } else {
        row[field.name] = "";
      }
    } else if (field.type === "boolean") {
      let fieldValue = fieldValueOnTable ?? false;

      // Handle string values like "TRUE", "YES", etc.
      if (typeof fieldValue === "string") {
        const upperCaseValue = fieldValue.toUpperCase();
        fieldValue = upperCaseValue === "TRUE" || upperCaseValue === "YES";
      }

      row[field.name] = fieldValue ? true : false;
    } else if (field.type === "number") {
      const fieldValue = fieldValueOnTable;
      let exportValue: number | null = null;
      if (fieldValue !== null && fieldValue !== undefined && fieldValue !== "") {
        const parsed = Number(fieldValue);
        if (!isNaN(parsed)) {
          exportValue = parsed;
          isEmpty = false;
        }
      }
      row[field.name] = exportValue;
    } else {
      const fieldValue = fieldValueOnTable ?? field?.defaultValue;
      if (fieldValue || fieldValue === 0) {
        isEmpty = false;
      }
      let exportValue = fieldValue;
      if (Array.isArray(fieldValue) && field.representedType === "array") {
        exportValue = fieldValue;
      } else if (!Array.isArray(fieldValue) && field.representedType === "array") {
        exportValue = fieldValue ? [fieldValue] : field.defaultValue;
      } else {
        exportValue = fieldValue.toString() || field.defaultValue;
      }
      row[field.name] = exportValue;
    }
    return [row, isEmpty];
  }

  if (field.type === "object") {
    row[field.name] = {
      "@context": contextUrl,
      "@type": field.objectType,
    };

    for (const property of field.properties || []) {
      // Call the function recursively
      const [newRow, newIsEmpty] = await getObjectFieldsRecursively(
        headers,
        values,
        property,
        row[field.name],
        isEmpty
      );
      row[field.name] = { ...row[field.name], ...newRow };
      isEmpty = newIsEmpty;
    }
  }

  return [row, isEmpty];
}
/* eslint-enable no-param-reassign */
