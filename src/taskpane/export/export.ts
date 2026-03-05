import moment from "moment-timezone";
import { IntlShape } from "react-intl";
import { CodeList, getCodeListByTableName } from "../domain/fetchServer/getCodeLists";
import {
  getUnitDefinition,
  UNIT_DEFINITIONS,
  UNIT_IRI,
} from "../domain/fetchServer/getUnitsOfMeasure";
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
import { downloadJSONLD, formatMessageToString } from "../utils/utils";

function removeNamespacePrefixesFromExport(items: any[]): any[] {
  return items.map((item) => {
    if (!item || typeof item !== "object") return item;

    const cleaned: any = {};

    for (const [key, value] of Object.entries(item)) {
      let newKey = key;

      // Preserve JSON-LD keywords (@context, @id, @type, etc.)
      if (key.startsWith("@")) {
        cleaned[key] = value;
        continue;
      }

      // Remove namespace prefix from property names
      // e.g., "rdfs:label" => "label", "i72:value" => "value"
      if (key.includes(":")) {
        const parts = key.split(":");
        if (parts.length === 2 && parts[0].length > 0 && parts[1].length > 0) {
          newKey = parts[1];
        }
      }

      // Recursively clean nested objects
      if (value && typeof value === "object" && !Array.isArray(value)) {
        const nested = removeNamespacePrefixesFromExport([value])[0];
        delete nested["@context"];
        cleaned[newKey] = nested;
      } else if (Array.isArray(value)) {
        // For arrays, check if items are objects that need cleaning
        const hasObjects = value.some((v) => v && typeof v === "object");
        cleaned[newKey] = hasObjects ? removeNamespacePrefixesFromExport(value) : value;
      } else {
        cleaned[newKey] = value;
      }
    }

    return cleaned;
  });
}
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
          formatMessageToString(intl, {
            id: "generics.error",
            defaultMessage: "Error",
          }),
          formatMessageToString(
            intl,
            {
              id: "export.messages.error.missingTable",
              defaultMessage:
                "Table <b>{tableName}</b> is missing. Please create the tables first.",
            },
            { tableName: key, b: (str: string) => `<b>${str}</b>` }
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
              // Skip @type field as it might not exist in the table
              if (fieldName === "@type") continue;

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
                formatMessageToString(
                  intl,
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
              if (key === "@id" || key === "@type") return true;

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
              formatMessageToString(
                intl,
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

        // Extract and set the @id
        const idColIndex = headers.indexOf("@id");
        if (idColIndex !== -1 && recordValues[idColIndex]) {
          row["@id"] = recordValues[idColIndex].toString();
        }

        let isEmpty = true; // Flag to check if the row is empty

        // Process all fields in one pass using the cached headers
        const [processedRow, rowIsEmpty] = await processRecord(
          tableFields,
          headers,
          recordValues,
          row,
          isEmpty
        );

        // Only add non-empty rows with valid @id
        if (!rowIsEmpty && processedRow["@id"]) {
          data.push(processedRow);
        }
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
      const typeVal = item["@type"];
      const types = Array.isArray(typeVal) ? typeVal : [typeVal];
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
      const typeVal = item["@type"];
      const types = Array.isArray(typeVal) ? typeVal : [typeVal];
      if (types.includes("cids:IndicatorReport")) {
        const indicatorId = item["forIndicator"]; // link field name assumption
        const valueObj = item["i72:value"] as Record<string, any> | undefined;
        if (valueObj && !valueObj["i72:unit_of_measure"]) {
          const fallback =
            (typeof indicatorId === "string" && indicatorUnitById[indicatorId]) ||
            UNIT_IRI.UNSPECIFIED;
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
            if (
              typeof val === "string" &&
              val.startsWith("https://ontology.commonapproach.org/cids#")
            ) {
              queue.push(val);
            }
          }
        }
      } catch {
        /* ignore */
      }
    }

    const { errors, warnings } = await validate(data, "export", intl);

    // Check for unexported fields and empty tables using new helper functions
    const unexportedFieldWarnings = await checkForUnexportedFields(context, tables, fullMap, intl);
    const emptyTableWarnings = await checkForEmptyTables(context, tables, fullMap, intl);

    // Include the code list warnings in the warnings
    const allWarnings = [
      ...unexportedFieldWarnings,
      ...warnings,
      ...emptyTableWarnings,
      ...changeOnDefaultCodeListsWarning,
    ]
      .filter(Boolean)
      .join("<hr/>");

    if (errors.length > 0) {
      setDialogContent(
        formatMessageToString(intl, {
          id: "generics.error",
          defaultMessage: "Error",
        }),
        errors.map((item) => `<p>${item}</p>`).join("")
      );
      return;
    }

    // Always deep-clean just before potentially showing warnings (but keep original for reference)
    const cleanedData = deepCleanExportObjects(data);
    const finalData = removeNamespacePrefixesFromExport(cleanedData);

    if (allWarnings.length > 0) {
      setDialogContent(
        formatMessageToString(intl, {
          id: "generics.warning",
          defaultMessage: "Warning",
        }),
        allWarnings,
        () => {
          setDialogContent(
            formatMessageToString(intl, {
              id: "generics.warning",
              defaultMessage: "Warning",
            }),
            formatMessageToString(intl, {
              id: "export.messages.warning.continue",
              defaultMessage: "<p>Do you want to export anyway?</p>",
            }),
            () => {
              downloadJSONLD(finalData, `${getFileName(orgName)}.json`);
              setDialogContent(
                formatMessageToString(intl, {
                  id: "generics.success",
                  defaultMessage: "Success",
                }),
                formatMessageToString(intl, {
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
    downloadJSONLD(finalData, `${getFileName(orgName)}.json`);
    setDialogContent(
      formatMessageToString(intl, {
        id: "generics.success",
        defaultMessage: "Success",
      }),
      formatMessageToString(intl, {
        id: "export.messages.success",
        defaultMessage: "Data exported successfully!",
      })
    );
  });
}

async function checkForEmptyTables(
  context: Excel.RequestContext,
  tables: Excel.Table[],
  fullMap: any,
  intl: IntlShape
): Promise<string[]> {
  const warnings: string[] = [];
  const relevantTables = tables.filter((table) => Object.keys(fullMap).includes(table.name));

  // Queue loads for all data ranges
  const tableData = relevantTables.map((table) => {
    const dataRange = table.getDataBodyRange();
    if (dataRange) {
      dataRange.load("rowCount");
    }
    return { table, dataRange };
  });

  // Execute a single sync for all loads
  await context.sync();

  // Check each table
  for (const { table, dataRange } of tableData) {
    // If dataRange doesn't exist OR rowCount is 0, the table is empty
    if (!dataRange || dataRange.rowCount === 0) {
      warnings.push(
        formatMessageToString(
          intl,
          {
            id: "export.messages.warning.emptyTable",
            defaultMessage: "Table <b>{tableName}</b> is empty",
          },
          {
            tableName: table.name,
            b: (str: string) => `<b>${str}</b>`,
          }
        )
      );
    }
  }

  return warnings;
}

// Improved function to check for unexported fields
async function checkForUnexportedFields(
  context: Excel.RequestContext,
  tables: Excel.Table[],
  fullMap: any,
  intl: IntlShape
): Promise<string[]> {
  const warnings: string[] = [];
  const relevantTables = tables.filter((table) => Object.keys(fullMap).includes(table.name));

  // Queue loads for all header ranges
  const tableData = relevantTables.map((table) => {
    const headerRange = table.getHeaderRowRange();
    headerRange.load("values");

    const cid = createInstance(table.name as ModelType | SFFModelType);
    const internalFields = cid.getAllFields().map((item) => item.displayName || item.name);

    return {
      table,
      headerRange,
      internalFields,
    };
  });

  // Execute a single sync for all loads
  await context.sync();

  // Check each table for unexported fields
  for (const { table, headerRange, internalFields } of tableData) {
    const headers = headerRange.values[0];
    for (const field of headers) {
      // Skip if it's a table name or in ignoredFields
      if (
        Object.keys(fullMap).includes(field) ||
        (ignoredFields as any)[table.name]?.includes(field)
      ) {
        continue;
      }

      // Warn if field is not in the model
      if (!internalFields.includes(field)) {
        warnings.push(
          formatMessageToString(
            intl,
            {
              id: Object.keys(map).includes(table.name)
                ? "export.messages.warning.fieldWillNotBeExported"
                : "export.messages.warning.notExported",
              defaultMessage:
                "Field <b>{fieldName}</b> in table <b>{tableName}</b> is inconsistent with the Basic Tier of the Common Impact Data Standard. This field will not be exported.",
            },
            { fieldName: field, tableName: table.name, b: (str: string) => `<b>${str}</b>` }
          ) as string
        );
      }
    }
  }

  return warnings;
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
        // Preserve user-entered value even if not in predefined options
        row[field.name] = field.representedType === "array" ? [fieldValue] : fieldValue;
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
      let optionFields: { id: string; name: string }[] = [];
      if (field.getOptionsAsync) {
        const options = await field.getOptionsAsync();
        optionFields = options.filter((opt) => valuesArray.includes(opt.name));
      } else {
        optionFields = field.selectOptions?.filter((opt) => valuesArray.includes(opt.name)) || [];
      }
      // Get IDs for recognized options
      const recognizedOptionIds = optionFields.map((opt) => opt.id);
      // Preserve unrecognized option names (custom values)
      const unrecognizedOptionNames = valuesArray.filter(
        (val) => !optionFields.some((opt) => opt.name === val)
      );
      const combinedValues = [...recognizedOptionIds, ...unrecognizedOptionNames];
      row[field.name] =
        field.representedType === "array" ? combinedValues : combinedValues.join(", ");
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
      const fieldValue = value ?? field.defaultValue;
      if (fieldValue || fieldValue === 0) {
        isEmpty = false;
      }
      let exportValue = fieldValue;
      if (Array.isArray(fieldValue) && field.representedType === "array") {
        exportValue = fieldValue;
      } else if (!Array.isArray(fieldValue) && field.representedType === "array") {
        // Handle comma-separated values for array fields (e.g., @type with multiple types)
        if (typeof fieldValue === "string" && fieldValue.includes(",")) {
          exportValue = fieldValue.split(",").map((s) => s.trim());
        } else {
          exportValue = fieldValue ? [fieldValue] : field.defaultValue;
        }
      } else {
        exportValue = fieldValue ? fieldValue.toString() : field.defaultValue;
      }
      row[field.name] = exportValue as any;
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
      const arr = value
        .map((v) => clean(v))
        .filter((v) => {
          if (v === null || v === undefined) return false;
          if (Array.isArray(v) && v.length === 0) return false;
          if (typeof v === "object" && !Array.isArray(v) && Object.keys(v).length === 0)
            return false;
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
          if (typeof v === "object" && !Array.isArray(v) && Object.keys(v).length === 0)
            return false;
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
