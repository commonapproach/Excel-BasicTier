import moment from "moment-timezone";
import { IntlShape } from "react-intl";
import { CodeList, getCodeListByTableName } from "../domain/fetchServer/getCodeLists";
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

    for (const table of tables) {
      // If the table is not in the map, skip it
      if (!Object.keys(fullMap).includes(table.name)) {
        continue;
      }

      // Get the records from the table
      const tableRange = table.getRange();
      const tableHeaderRange = tableRange.getRow(0);
      tableHeaderRange.load("values");
      table.load("values, rows");
      await context.sync();
      const records = table.rows.items;
      const idColumnIndex = tableHeaderRange.values[0].indexOf("@id");

      let codeList: CodeList[] | null = null;
      if (predefinedCodeLists.includes(table.name)) {
        codeList = await getCodeListByTableName(table.name);
      }

      const cid: BaseModel = createInstance(table.name as ModelType | SFFModelType);
      for (const record of records) {
        record.load("values");
        await context.sync();

        // Get the headers to map column indices to field names
        const headers = tableHeaderRange.values[0];

        // Skip records that are defined in the common approach code lists
        if (
          codeList &&
          idColumnIndex !== -1 &&
          codeList.find((item) => item["@id"] === record.values[0][idColumnIndex])
        ) {
          // Check if the record has changes compared to the code list item
          const recordId = record.values[0][idColumnIndex];
          const existingItem = codeList.find((item) => item["@id"] === recordId);

          if (existingItem) {
            let hasChanges = false;
            for (const fieldName of Object.keys(existingItem)) {
              const fieldIndex = headers.indexOf(fieldName);
              if (fieldIndex !== -1) {
                const recordValue = record.values[0][fieldIndex];
                const existingValue = (existingItem as Record<string, any>)[fieldName];

                // Compare values (simple comparison, might need refinement)
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
          const recordValues = record.values[0];

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

        let row: TableInterface = {
          "@context": contextUrl,
          "@type": `cids:${table.name}`,
          "@id": "",
        };

        let isEmpty = true; // Flag to check if the row is empty

        for (const field of cid.getTopLevelFields()) {
          const columnIndex = tableHeaderRange.values[0].indexOf(field.displayName || field.name);
          const value: any = record.values[0][columnIndex];
          if (field.type === "link") {
            if (field.representedType === "array") {
              const fieldValue = value ?? field?.defaultValue;
              if (fieldValue && fieldValue.length > 0) {
                isEmpty = false;
              }
              row[field.name] =
                typeof fieldValue === "string"
                  ? fieldValue.split(", ").filter((v) => v !== "" && v !== null && v !== undefined)
                  : (fieldValue as string[]).filter(
                      (v) => v !== "" && v !== null && v !== undefined
                    );
            } else if (field.representedType === "string") {
              const fieldValue = value ?? field?.defaultValue;
              if (fieldValue) {
                isEmpty = false;
              }
              row[field.name] = Array.isArray(fieldValue) ? fieldValue[0] : fieldValue;
            }
          } else if (field.type === "object") {
            const [newRow, newIsEmpty] = getObjectFieldsRecursively(
              tableHeaderRange.values[0],
              record.values[0],
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
              row[field.name] =
                field.representedType === "array" ? [optionField.id] : optionField.id;
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
              optionFields =
                field.selectOptions?.filter((opt) => valuesArray.includes(opt.name)) || [];
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
            const fieldValue = value ?? false;
            row[field.name] = fieldValue ? true : false;
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
            row[field.name] = exportValue;
          }
        }
        if (!isEmpty) {
          data.push(row);
        }
      }
    }

    const { errors, warnings } = await validate(data, "export", intl);

    const noExportingFields = await checkForNotExportedFields(intl, context);
    const emptyTableWarning = await checkForEmptyTables(intl, context);

    // Include the code list warnings in the warnings
    let codeListWarnings: string[] = [];
    if (changeOnDefaultCodeListsWarning.length > 0) {
      codeListWarnings = changeOnDefaultCodeListsWarning;
    }

    const allWarnings = [
      ...noExportingFields,
      ...warnings,
      ...emptyTableWarning,
      ...codeListWarnings,
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
              downloadJSONLD(data, `${getFileName(orgName)}.json`);
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
    downloadJSONLD(data, `${getFileName(orgName)}.json`);
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

async function checkForNotExportedFields(intl: IntlShape, context: Excel.RequestContext) {
  const workbook = context.workbook;
  workbook.load("tables");
  await context.sync();
  const tables = workbook.tables.items;
  const fullMap = { ...map, ...mapSFFModel };

  let warnings: string[] = [];
  for (const table of tables) {
    if (!Object.keys(fullMap).includes(table.name)) {
      continue;
    }
    const cid = createInstance(table.name as ModelType | SFFModelType);
    const internalFields = cid.getAllFields().map((item) => item.displayName || item.name);

    const tableHeaderRange = table.getHeaderRowRange();
    tableHeaderRange.load("values");
    await context.sync();
    const tableHeadersValues = tableHeaderRange.values[0];

    for (const field of tableHeadersValues) {
      if (
        Object.keys(fullMap).includes(field) ||
        (ignoredFields as any)[table.name]?.includes(field)
      ) {
        continue;
      }
      if (!internalFields.includes(field)) {
        warnings.push(
          intl.formatMessage(
            {
              id: "export.messages.warning.fieldWillNotBeExported",
              defaultMessage:
                "Field <b>{fieldName}</b> on table <b>{tableName}</b> will not be exported",
            },
            { fieldName: field, tableName: table.name, b: (str: string) => `<b>${str}</b>` }
          ) as string
        );
      }
    }
  }
  return warnings;
}

async function checkForEmptyTables(intl: IntlShape, context: Excel.RequestContext) {
  const workbook = context.workbook;
  workbook.load("tables");
  await context.sync();
  const tables = workbook.tables.items;
  const fullMap = { ...map, ...mapSFFModel };

  let warnings: string[] = [];
  for (const table of tables) {
    if (!Object.keys(fullMap).includes(table.name)) {
      continue;
    }

    const tableDataRange = table.getDataBodyRange();
    tableDataRange.load("values");
    await context.sync();
    const tableData = tableDataRange.values;

    let isEmpty = true;
    for (const row of tableData) {
      for (const cell of row) {
        if (cell) isEmpty = false;
      }
    }

    if (isEmpty) {
      warnings.push(
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
  return warnings;
}

/* eslint-disable no-param-reassign */
function getObjectFieldsRecursively(
  headers: string[],
  value: any,
  field: FieldType,
  row: any,
  isEmpty: boolean
) {
  if (field.type !== "object") {
    const columnIndex = headers.indexOf(field.displayName || field.name);
    const fieldValueOnTable = value[columnIndex];
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
      row[field.name] = fieldValueOnTable ? true : false;
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
      const [newRow, newIsEmpty] = getObjectFieldsRecursively(
        headers,
        value,
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
