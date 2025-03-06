import { IntlShape } from "react-intl";
import { TableInterface } from "../domain/interfaces/table.interface";
import {
  ModelType,
  SFFModelType,
  createInstance,
  ignoredFields,
  map,
  mapSFFModel,
} from "../domain/models";
import { FieldType } from "../domain/models/Base";
import { validate } from "../domain/validation/validator";
import { createSFFModuleSheetsAndTables, createSheetsAndTables } from "../taskpane";
import { parseJsonLd } from "../utils/utils";

/* global Excel */
export async function importData(
  intl: IntlShape,
  jsonData: any,
  setDialogContent: (header: string, content: string, nextCallBack?: Function) => void,
  setIsLoading: (isLoading: boolean) => void
) {
  await Excel.run(async (context) => {
    if (validateIfEmptyFile(jsonData)) {
      setDialogContent(
        intl.formatMessage({
          id: "generics.error",
          defaultMessage: "Error",
        }),
        intl.formatMessage({
          id: "import.messages.error.emptyOrNotArray",
          defaultMessage: "Table data is empty or not an array",
        })
      );
      return;
    }

    if (!doAllRecordsHaveId(jsonData)) {
      setDialogContent(
        intl.formatMessage({
          id: "generics.error",
          defaultMessage: "Error",
        }),
        intl.formatMessage({
          id: "import.messages.error.missingId",
          defaultMessage: "All records must have an <b>@id</b> property.",
        })
      );
      return;
    }

    // eslint-disable-next-line no-param-reassign
    jsonData = await parseJsonLd(jsonData);

    // Remove duplicated links
    // eslint-disable-next-line no-param-reassign
    jsonData = removeDuplicatedLinks(jsonData);

    let allErrors = "";
    let allWarnings = "";

    // Check if json data is a valid json array
    if (!Array.isArray(jsonData)) {
      setDialogContent(
        intl.formatMessage({
          id: "generics.error",
          defaultMessage: "Error",
        }),
        intl.formatMessage({
          id: "import.messages.error.invalidJson",
          defaultMessage: "Invalid JSON data, please check the data and try again.",
        })
      );
      return;
    }

    // Transform object field if it's in the wrong format
    // eslint-disable-next-line no-param-reassign
    jsonData = transformObjectFieldIfWrongFormat(jsonData);

    // Validate JSON
    let { errors, warnings } = await validate(jsonData, "import", intl);

    warnings = [...warnings, ...warnIfUnrecognizedFieldsWillBeIgnored(jsonData, intl)];

    allErrors = errors.join("<hr/>");
    allWarnings = warnings.join("<hr/>");

    if (allErrors.length > 0) {
      setDialogContent(
        intl.formatMessage({
          id: "generics.error",
          defaultMessage: "Error",
        }),
        allErrors
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
              id: "import.messages.warning.continue",
              defaultMessage: "<p>Do you want to import anyway?</p>",
            }),
            async () => {
              try {
                setIsLoading(true);
                await importFileData(intl, context, jsonData, setDialogContent);
              } finally {
                setIsLoading(false);
              }
            }
          );
        }
      );
    } else {
      await importFileData(intl, context, jsonData, setDialogContent);
    }
  });
}

async function importFileData(
  intl: IntlShape,
  context: Excel.RequestContext,
  jsonData: any,
  setDialogContent: any
) {
  setDialogContent(
    intl.formatMessage({
      id: "import.messages.wait",
      defaultMessage: "Wait a moment...",
    }),
    intl.formatMessage({
      id: "import.messages.importing",
      defaultMessage: "Importing data...",
    })
  );
  try {
    // Ignore types/classes that are not recognized
    const fullMap = { ...map, ...mapSFFModel };
    const filteredItems = Array.isArray(jsonData)
      ? jsonData.filter((data: any) => Object.keys(fullMap).includes(data["@type"].split(":")[1]))
      : jsonData;
    await importByData(intl, context, filteredItems);
  } catch (error: any) {
    setDialogContent(
      intl.formatMessage({
        id: "generics.error",
        defaultMessage: "Error",
      }),
      error.message ||
        intl.formatMessage({ id: "generics.error.message", defaultMessage: "Something went wrong" })
    );
    return;
  }
  setDialogContent(
    intl.formatMessage({
      id: "generics.success",
      defaultMessage: "Success",
    }),
    intl.formatMessage({
      id: "import.messages.success",
      defaultMessage: "Your data has been successfully imported.",
    })
  );
}

async function importByData(intl: IntlShape, context: Excel.RequestContext, jsonData: any) {
  // Set active worksheet Waiting sheet
  const randomValue = Math.random().toString(36).substring(7);
  const waitingSheetName = `Waiting${randomValue}`;
  const worksheets = context.workbook.worksheets;
  worksheets.load("items");
  await context.sync();
  let waitingSheetExists = false;
  for (const worksheet of worksheets.items) {
    if (worksheet.name === waitingSheetName) {
      worksheet.activate();
      waitingSheetExists = true;
    }
  }
  if (!waitingSheetExists) {
    const waitingSheet = worksheets.add(waitingSheetName);
    waitingSheet.activate();
  }
  await context.sync();

  // Write message to A1 warning users to do not edit the workbook while importing
  const range = context.workbook.worksheets.getItem(waitingSheetName).getRange("A1");
  range.values = [
    [
      intl.formatMessage({
        id: "import.messages.workbook.waiting",
        defaultMessage: "Do not edit this workbook while importing data.",
      }),
    ],
  ];
  await context.sync();

  // Create Tables if they don't exist
  await createSheetsAndTables();

  // Check if data has any class from SFF module
  for (const data of jsonData) {
    if (
      !data["@type"] ||
      (!Object.keys(map).includes(data["@type"].split(":")[1]) &&
        !Object.keys(mapSFFModel).includes(data["@type"].split(":")[1]))
    ) {
      continue;
    } else if (Object.keys(mapSFFModel).includes(data["@type"].split(":")[1])) {
      await createSFFModuleSheetsAndTables();
      break;
    }
  }

  // Write Simple Records to Tables
  await writeTable(context, jsonData);

  // Write Linked Records to Tables
  await writeTableLinked(context, jsonData);

  // Resize the tables
  worksheets.load("items");
  await context.sync();
  for (const worksheet of worksheets.items) {
    const tables = worksheet.tables;
    tables.load("items");
    await context.sync();
    for (const table of tables.items) {
      table.getRange().format.autofitColumns();
      table.getRange().format.autofitRows();
    }
  }
  await context.sync();

  // Remove the waiting sheet
  const waitingSheet = context.workbook.worksheets.getItem(waitingSheetName);
  waitingSheet.delete();
  await context.sync();
}

async function writeTable(
  context: Excel.RequestContext,
  tableData: TableInterface[]
): Promise<void> {
  for (const data of tableData) {
    const tableName = data["@type"].split(":")[1];
    const recordId = data["@id"];
    const worksheet = context.workbook.worksheets.getItem(tableName);
    worksheet.load("tables");
    context.trackedObjects.add(worksheet);
    await context.sync();
    const table = worksheet.tables.getItem(tableName);
    const tableRange = table.getRange();
    const tableHeaderRange = table.getHeaderRowRange();
    tableRange.load("values");
    tableHeaderRange.load("values");
    context.trackedObjects.add(tableRange);
    context.trackedObjects.add(tableHeaderRange);
    await context.sync();
    const idColumnIndex = tableHeaderRange.values[0].indexOf("@id");
    const idColumn = tableRange.getColumn(idColumnIndex);
    idColumn.load("values");
    context.trackedObjects.add(idColumn);
    await context.sync();
    const idColumnValues = idColumn.values;

    // Create the record
    let record: { [key: string]: unknown } = {};
    Object.entries(data).forEach(async ([key, value]) => {
      if (key === "@type" || key === "@context" || !checkIfFieldIsRecognized(tableName, key)) {
        return;
      }

      let cid;
      if (Object.keys(map).includes(tableName)) {
        cid = new map[tableName as ModelType]();
      } else {
        cid = new mapSFFModel[tableName as SFFModelType]();
      }

      for (const field of cid.getAllFields()) {
        if (field.name.includes(":") && field.name.split(":")[1] === key) {
          // eslint-disable-next-line no-param-reassign
          key = field.name;
          break;
        }
      }

      if (cid.getFieldByName(key)?.type !== "link" && cid.getFieldByName(key)?.type) {
        if (cid.getFieldByName(key)?.type === "object") {
          record = findLastFieldValueForNestedFields(data, cid.getFieldByName(key), record);
        } else {
          const field = cid.getFieldByName(key);
          const fieldName = field.displayName || field.name;
          let newValue: any = value;
          if (newValue && (field.type === "select" || field.type === "multiselect")) {
            let options: { id: string; name: string }[] = [];
            if (field.getOptionsAsync) {
              options = await field.getOptionsAsync();
            } else {
              options = field.selectOptions as { id: string; name: string }[];
            }
            if (
              field.type === "select" &&
              options.find((opt) => opt.id === (Array.isArray(newValue) ? newValue[0] : newValue))
            ) {
              const optionField = options.find(
                (opt) => opt.id === (Array.isArray(newValue) ? newValue[0] : newValue)
              );
              if (optionField) {
                newValue = optionField.name;
              } else {
                newValue = null;
              }
            } else if (field.type === "multiselect") {
              newValue = (newValue as string[]).map((val) => {
                const optionField = options.find((opt) => opt.id === val);
                if (optionField) {
                  return optionField.name;
                }
                return null;
              });
              // Remove null values
              newValue = newValue.filter((val: any) => val);
              if (newValue.length === 0) {
                newValue = null;
              } else {
                newValue = newValue.join(", ");
              }
            } else {
              newValue = null;
            }
          }
          if (field.type === "boolean") {
            if (newValue && (newValue === true || (newValue as string).toLowerCase() === "true")) {
              newValue = true;
            } else {
              newValue = false;
            }
          }
          if (
            field.type !== "boolean" &&
            field.type !== "select" &&
            field.type !== "multiselect" &&
            newValue
          ) {
            newValue = newValue.toString();
          }
          record[fieldName] = newValue;
        }
      }
    });

    // Add or Update the record on the table
    // check if the record already exists
    const idColumnValue = idColumnValues.map((item) => item[0].toString());
    const idIndex = idColumnValue.indexOf(recordId);
    let row: Excel.Range;
    if (idIndex !== -1) {
      row = tableRange.getRow(idIndex);
    } else {
      // Add the record
      // Get first row with empty id
      const emptyIdIndex = idColumnValue.indexOf("");
      row = tableRange.getRow(emptyIdIndex);
    }
    row.load("values");
    context.trackedObjects.add(row);
    await context.sync();

    // Update the record
    for (const [key, value] of Object.entries(record)) {
      const columnIndex = tableHeaderRange.values[0].indexOf(key);
      row.values = row.values || [];
      row.values[0][columnIndex] = value;
    }

    context.trackedObjects.remove(row);
    context.trackedObjects.remove(idColumn);
    context.trackedObjects.remove(tableRange);
    context.trackedObjects.remove(tableHeaderRange);
    context.trackedObjects.remove(worksheet);
    await context.sync();
  }
}

async function writeTableLinked(
  context: Excel.RequestContext,
  tableData: TableInterface[]
): Promise<void> {
  for (const data of tableData) {
    const tableName = data["@type"].split(":")[1];
    const recordId = data["@id"];
    const worksheet = context.workbook.worksheets.getItem(tableName);
    worksheet.load("tables");
    context.trackedObjects.add(worksheet);
    await context.sync();
    const table = worksheet.tables.getItem(tableName);
    const tableRange = table.getRange();
    const tableHeaderRange = table.getHeaderRowRange();
    tableRange.load("values");
    tableHeaderRange.load("values");
    context.trackedObjects.add(tableRange);
    context.trackedObjects.add(tableHeaderRange);
    await context.sync();
    const idColumnIndex = tableHeaderRange.values[0].indexOf("@id");
    const idColumn = tableRange.getColumn(idColumnIndex);
    idColumn.load("values");
    context.trackedObjects.add(idColumn);
    await context.sync();
    const idColumnValues = idColumn.values;

    // Create the record
    let record: { [key: string]: unknown } = {};
    for (let [key, value] of Object.entries(data)) {
      const cid = createInstance(tableName as ModelType | SFFModelType);

      if (key !== "@type" && key !== "@context" && checkIfFieldIsRecognized(tableName, key)) {
        const field = cid.getFieldByName(key);
        if (field) {
          record = findLinkFieldsRecursively(tableName, field, data, record);
        }
      }
    }

    // Add or Update the record on the table
    // check if the record already exists
    const idColumnValue = idColumnValues.map((item) => item[0].toString());
    let idIndex = idColumnValue.indexOf(recordId);
    let row: Excel.Range;
    if (idIndex !== -1) {
      row = tableRange.getRow(idIndex);
    } else {
      // Add the record
      // Get first row with empty id
      idIndex = idColumnValue.indexOf("");
      row = tableRange.getRow(idIndex);
    }
    row.load("values");
    context.trackedObjects.add(row);
    await context.sync();
    const rowValues = row.values[0];

    // Update the record
    for (let [key, value] of Object.entries(record)) {
      const columnIndex = tableHeaderRange.values[0].indexOf(key);
      value = [
        ...new Set([
          ...((rowValues[columnIndex] as string).split(", ") || []),
          ...((value as string[]) || []),
        ]),
      ].filter((v) => v !== null && v !== undefined && v !== "");
      row.getCell(0, columnIndex).values = [[(value as string[]).join(", ")]];
    }

    context.trackedObjects.remove(row);
    await context.sync();

    // Update the linked tables
    for (const [key, value] of Object.entries(record)) {
      const cid = createInstance(tableName as ModelType | SFFModelType);

      let field: FieldType | null = null;
      try {
        field = cid.getFieldByName(key);
      } catch (_) {
        continue;
      }

      if (
        field.link &&
        field.link.table &&
        tableName !== field.link.table.className &&
        (!(field.link.table.className in ignoredFields) ||
          !(ignoredFields as any)[field.link.table.className].includes(field.link.field))
      ) {
        await updateLinkedTablesFields(
          context,
          recordId,
          (value as string[]) || [],
          field.link.table.className,
          field.link.field
        );
      }
    }

    context.trackedObjects.remove(idColumn);
    context.trackedObjects.remove(tableRange);
    context.trackedObjects.remove(tableHeaderRange);
    context.trackedObjects.remove(worksheet);
    await context.sync();
  }
}

/* eslint-disable no-param-reassign */
function handleLinkFields(value: any) {
  if (!value) return; // Skip if the value is empty

  if (!Array.isArray(value)) value = [value];

  // remove duplicates from value
  value = [...new Set(value as string[])];

  if (value.length === 0) {
    return;
  }

  return value;
}
/* eslint-enable no-param-reassign */

function findLinkFieldsRecursively(tableName: string, field: FieldType, data: any, record: any) {
  if (field?.type === "object" && field.properties) {
    for (const prop of field.properties) {
      findLinkFieldsRecursively(tableName, prop, data[field.name], record);
    }
  } else {
    if (field.type === "link") {
      // eslint-disable-next-line no-param-reassign
      record[field.displayName || field.name] = handleLinkFields(data[field.name]);
    }
  }
  return record;
}

async function updateLinkedTablesFields(
  context: Excel.RequestContext,
  currentFieldId: string,
  currentFiledValues: string[],
  relatedTableName: string,
  relatedFieldName: string
) {
  const worksheet = context.workbook.worksheets.getItem(relatedTableName);
  worksheet.load("tables");
  context.trackedObjects.add(worksheet);
  await context.sync();
  const table = worksheet.tables.getItem(relatedTableName);
  const tableHeadersRange = table.getHeaderRowRange();
  const tableRange = table.getRange();
  tableHeadersRange.load("values");
  tableRange.load("values");
  context.trackedObjects.add(tableHeadersRange);
  context.trackedObjects.add(tableRange);
  await context.sync();
  const relatedFieldIndex = tableHeadersRange.values[0].indexOf(relatedFieldName);
  const idColumnIndex = tableHeadersRange.values[0].indexOf("@id");
  const relatedFieldColumn = tableRange.getColumn(relatedFieldIndex);
  const idColumn = tableRange.getColumn(idColumnIndex);
  relatedFieldColumn.load("values");
  idColumn.load("values");
  context.trackedObjects.add(relatedFieldColumn);
  context.trackedObjects.add(idColumn);
  await context.sync();
  const relatedFieldValues = relatedFieldColumn.values;
  const idColumnValues = idColumn.values;

  // For each cell in the related field column, check if the id is in the array
  for (let i = 0; i < relatedFieldValues.length; i++) {
    const idColumnValue = idColumnValues[i][0].toString();
    const relatedFieldValue = relatedFieldValues[i][0].toString();
    const relatedFieldValueArray: string[] = relatedFieldValue.split(", ");

    if (!idColumnValue) {
      continue;
    }

    if (
      currentFiledValues.includes(idColumnValue) &&
      !relatedFieldValueArray.includes(currentFieldId)
    ) {
      // Add the id to the array
      relatedFieldColumn.getCell(i, 0).values = [
        [
          relatedFieldValueArray
            .concat(currentFieldId)
            .filter((v) => v !== null && v !== undefined && v !== "")
            .join(", "),
        ],
      ];
    }
  }

  context.trackedObjects.remove(relatedFieldColumn);
  context.trackedObjects.remove(idColumn);
  context.trackedObjects.remove(tableRange);
  context.trackedObjects.remove(tableHeadersRange);
  context.trackedObjects.remove(worksheet);
  await context.sync();
}

function removeDuplicatedLinks(jsonData: any) {
  for (const data of jsonData) {
    for (const [key, value] of Object.entries(data)) {
      if (Array.isArray(value)) {
        data[key] = [...new Set(value)];
      } else if (value && typeof value === "object") {
        removeDuplicatedLinksRecursively(value);
      }
    }
  }
  return jsonData;
}

function removeDuplicatedLinksRecursively(data: any) {
  for (const [key, value] of Object.entries(data)) {
    if (Array.isArray(value)) {
      // eslint-disable-next-line no-param-reassign
      data[key] = [...new Set(value)];
    } else if (typeof value === "object") {
      removeDuplicatedLinksRecursively(value);
    }
  }
}

function validateIfEmptyFile(tableData: TableInterface[]) {
  if (!Array.isArray(tableData) || tableData.length === 0) {
    return true;
  }
  return false;
}

function doAllRecordsHaveId(tableData: TableInterface[]) {
  for (const data of tableData) {
    if (data["@id"] === undefined) {
      return false;
    }
  }
  return true;
}

function warnIfUnrecognizedFieldsWillBeIgnored(tableData: TableInterface[], intl: IntlShape) {
  const warnings = [];
  const classesSet = new Set();
  for (const data of tableData) {
    if (
      !data["@type"] ||
      (!Object.keys(map).includes(data["@type"].split(":")[1]) &&
        !Object.keys(mapSFFModel).includes(data["@type"].split(":")[1]))
    ) {
      continue;
    }

    const tableName = data["@type"].split(":")[1];
    if (classesSet.has(tableName)) {
      continue;
    }

    for (const key in data) {
      if (key !== "@type" && key !== "@context" && !checkIfFieldIsRecognized(tableName, key)) {
        warnings.push(
          `${intl.formatMessage(
            {
              id: "import.messages.warning.unrecognizedField",
              defaultMessage:
                "Table <b>{tableName}</b> has unrecognized field <b>{fieldName}</b>. This field will be ignored.",
            },
            { tableName, fieldName: key, b: (str) => `<b>${str}</b>` }
          )}`
        );
        classesSet.add(tableName);
      }
    }
  }
  return warnings;
}

function checkIfFieldIsRecognized(tableName: string, fieldName: string) {
  const cid = createInstance(tableName as ModelType | SFFModelType);
  return cid
    .getAllFields()
    .reduce((acc: string[], field) => {
      acc.push(field.name);
      if (field.name.includes(":")) {
        acc.push(field.name.split(":")[1]);
      }
      return acc;
    }, [])
    .includes(fieldName);
}

function findLastFieldValueForNestedFields(data: any, field: FieldType, record: any) {
  if (field?.type === "object" && field.properties) {
    for (const prop of field.properties) {
      let dataPropName;
      if (field.name.includes(":") && Object.keys(data).includes(field.name.split(":")[1])) {
        dataPropName = field.name.split(":")[1];
      } else {
        dataPropName = field.name;
      }
      const recordData = data[dataPropName];
      findLastFieldValueForNestedFields(recordData, prop, record);
    }
  } else if (data && typeof data === "object" && !Array.isArray(data)) {
    let recordData;
    if (field.name.includes(":") && Object.keys(data).includes(field.name.split(":")[1])) {
      recordData = data[field.name.split(":")[1]];
    } else {
      recordData = data[field.name];
    }
    // eslint-disable-next-line no-param-reassign
    record[field.displayName || field.name] = recordData;
  } else {
    let value = data;
    if (value) {
      value = data.toString();
    }
    // eslint-disable-next-line no-param-reassign
    record[field.displayName || field.name] = value;
  }

  return record;
}

function transformObjectFieldIfWrongFormat(jsonData: TableInterface[]) {
  for (const data of jsonData) {
    for (const [key, value] of Object.entries(data)) {
      if (
        !data["@type"] ||
        (!Object.keys(map).includes(data["@type"].split(":")[1]) &&
          !Object.keys(mapSFFModel).includes(data["@type"].split(":")[1]))
      ) {
        continue;
      }

      if (
        key === "@type" ||
        key === "@context" ||
        key === "@id" ||
        !checkIfFieldIsRecognized(data["@type"].split(":")[1], key)
      ) {
        continue;
      }

      let cid;
      if (Object.keys(map).includes(data["@type"].split(":")[1])) {
        cid = new map[data["@type"].split(":")[1] as ModelType]();
      } else {
        cid = new mapSFFModel[data["@type"].split(":")[1] as SFFModelType]();
      }

      const field = cid.getFieldByName(key);
      if (field?.type === "object") {
        const fieldValue = handleNestedObjectFieldType(jsonData, field, value);
        if (fieldValue) {
          data[key] = fieldValue;
        }
      } else if (field?.type === "link" && typeof value === "object" && !Array.isArray(value)) {
        let id = value["@id"];
        if (!id) {
          id =
            data["@id"] +
            "/" +
            (value["@type"].includes(":") ? value["@type"].split(":")[1] : value["@type"]);
        }
        value["@id"] = id;
        jsonData.push(value);
        data[key] = id;
      }
    }
  }
  return jsonData;
}

function handleNestedObjectFieldType(data: TableInterface[], field: FieldType, value: any) {
  let fieldValue: any = null;
  if (field?.type === "object" && typeof value === "string") {
    fieldValue = data.find((d) => d["@id"] === value);
  } else if (field?.type === "object" && Array.isArray(value)) {
    fieldValue = data.find((d) => d["@id"] === value[0]);
  } else if (field?.type === "object" && typeof value === "object" && field.properties) {
    for (const prop of field.properties) {
      const newValue = handleNestedObjectFieldType(data, prop, value[prop.name]);
      if (newValue) {
        fieldValue = { [field.name]: newValue };
      }
    }
  }
  return fieldValue;
}
