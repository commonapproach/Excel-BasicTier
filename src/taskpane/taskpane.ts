import { dialogHandler } from "./context/DialogContext";
import { getCodeListByTableName } from "./domain/codeLists/getCodeLists";
import {
  createInstance,
  ignoredFields,
  map,
  mapSFFModel,
  ModelType,
  predefinedCodeLists,
  SFFModelType,
} from "./domain/models";
import {
  checkIfAllValuesExistInRelatedSheet,
  handleLinkedFieldsInRelatedSheet,
  updateRelatedFieldsValues,
} from "./helpers/handleLinkedFieldsOnOtherSheet";

const hiddenSheets = [
  "ProvinceTerritory",
  "OrganizationType",
  "Locality",
  "StreetType",
  "StreetDirection",
];

/* global Office */
Office.onReady(() => {
  // If needed, Office.js is ready to be called.

  // Add multi-select functionality if the user have all the standard tables
  if (Office.context.host === Office.HostType.Excel) {
    addMultiSelectHandlerToAllTables();
  }
});

/* global Excel console */
export async function createSheetsAndTables() {
  await Excel.run(async (context) => {
    try {
      const sheets = context.workbook.worksheets;

      // Create new sheets
      for (const sheetName of Object.keys(map)) {
        try {
          sheets.add(sheetName);
          await context.sync();
        } catch (error) {
          console.log("Error: " + error);
        }
      }
      await context.sync();

      // Create Hidden Sheets For Single Select Fields
      for (const hidden of ["StreetType", "StreetDirection"]) {
        try {
          const hiddenSheet = sheets.add(hidden);
          hiddenSheet.visibility = "Hidden";
          await context.sync();
        } catch (error) {
          console.log("Error: " + error);
        }
      }

      // Add tables to the sheets
      for (const sheetName of Object.keys(map)) {
        try {
          await createTable(context, sheetName as ModelType);
        } catch (error) {
          console.log("Error: " + error);
        }
      }

      // Add tables to the hidden sheets
      for (const hidden of ["StreetType", "StreetDirection"]) {
        try {
          await createHiddenTables(context, hidden);
        } catch (error) {
          console.log("Error: " + error);
        }
      }

      // Add multi-select functionality to link fields
      for (const sheetName of Object.keys(map)) {
        try {
          await addMultiSelectToLinkFieldsOnSheet(context, sheetName as ModelType);
        } catch (error) {
          console.log("Error: " + error);
        }
      }

      // Populate select lists
      await populateSelectLists();

      await context.sync();
    } catch (error) {
      console.log("Error: " + error);
    }
  });
}

export async function createSFFModuleSheetsAndTables() {
  await Excel.run(async (context) => {
    try {
      const sheets = context.workbook.worksheets;

      // Create new sheets
      for (const sheetName of Object.keys(mapSFFModel)) {
        try {
          sheets.add(sheetName);
          await context.sync();
        } catch (error) {
          console.log("Error: " + error);
        }
      }
      await context.sync();

      // Create Hidden Sheets For Single Select Fields
      for (const hidden of hiddenSheets) {
        try {
          const hiddenSheet = sheets.add(hidden);
          hiddenSheet.visibility = "Hidden";
          await context.sync();
        } catch (error) {
          console.log("Error: " + error);
        }
      }

      // Add tables to the sheets
      for (const sheetName of Object.keys(mapSFFModel)) {
        try {
          await createTable(context, sheetName as SFFModelType);
        } catch (error) {
          console.log("Error: " + error);
        }
      }

      // Add tables to the hidden sheets
      for (const hidden of hiddenSheets) {
        try {
          await createHiddenTables(context, hidden);
        } catch (error) {
          console.log("Error: " + error);
        }
      }

      // Add multi-select functionality to link fields
      for (const sheetName of Object.keys(mapSFFModel)) {
        try {
          await addMultiSelectToLinkFieldsOnSheet(context, sheetName as SFFModelType);
        } catch (error) {
          console.log("Error: " + error);
        }
      }

      // Populate code lists
      await populateCodeLists();

      await context.sync();
    } catch (error) {
      console.log("Error: " + error);
    }
  });
}

async function createTable(context: Excel.RequestContext, sheetName: ModelType | SFFModelType) {
  const sheet = context.workbook.worksheets.getItem(sheetName);

  const tableName = sheetName.toString();
  const fields = getTableFieldsNames(sheetName);

  // If header is '@id', replace it with "'@id" to avoid Excel error
  const headers = fields.map((field) => (field === "@id" ? "'@id" : field)); // eslint-disable-line quotes

  let table: Excel.Table | null;

  // Check if the table already exists
  try {
    table = sheet.tables.getItem(tableName);
    await context.sync();
  } catch (error) {
    table = null;
  }

  if (table) {
    // Check if the table has the correct headers
    const tableHeaders = table.getHeaderRowRange();
    tableHeaders.load("values");
    await context.sync();
    const tableHeadersValues = tableHeaders.values[0];

    // Check if the table has the correct headers
    const headersToAdd: string[] = [];
    for (let header of headers) {
      // eslint-disable-next-line quotes
      if (header === "'@id") {
        header = "@id";
      }
      if (!tableHeadersValues.includes(header)) {
        headersToAdd.push(header);
      }
    }

    // Expand the table to add the new headers
    if (headersToAdd.length > 0) {
      const newHeaders = tableHeadersValues.concat(headersToAdd);
      const headersRange = tableHeaders.getResizedRange(0, headersToAdd.length);
      headersRange.values = [newHeaders];
    }
  } else {
    // Add a table to the sheet
    const rangeAddress = `A1:${String.fromCharCode(65 + headers.length - 1)}1`;
    const tableRange = sheet.getRange(rangeAddress);
    tableRange.values = [headers];
    tableRange.format.autofitColumns();
    tableRange.format.autofitRows();

    // Resize the range to 1000 rows
    table = sheet.tables.add(tableRange.getResizedRange(1000, 0), true);
    table.name = tableName;
    table.showTotals = false;
    table.getRange().format.wrapText = true;
    table.getRange().format.verticalAlignment = "Center";
    table.getHeaderRowRange().format.columnWidth = 220;
    sheet.freezePanes.freezeRows(1);
  }

  await context.sync();
}

const getTableFieldsNames = (sheetName: ModelType | SFFModelType) => {
  const newClass = createInstance(sheetName);
  const fieldNames = [];
  const allFields = newClass.getAllFields();

  for (const field of allFields) {
    if (field.type === "object") {
      continue;
    }

    fieldNames.push(field.displayName || field.name);
  }

  return fieldNames;
};

async function createHiddenTables(context: Excel.RequestContext, tableName: string) {
  const sheet = context.workbook.worksheets.getItem(tableName);

  const headers = ["id", "name"];

  let table: Excel.Table | null;

  // Check if the table already exists
  try {
    table = sheet.tables.getItem(tableName);
    await context.sync();
  } catch (error) {
    table = null;
  }

  if (table) {
    return;
  } else {
    // Add a table to the sheet
    const rangeAddress = `A1:${String.fromCharCode(65 + headers.length - 1)}1`;
    const tableRange = sheet.getRange(rangeAddress);
    tableRange.values = [headers];
    tableRange.format.autofitColumns();
    tableRange.format.autofitRows();

    // Resize the range to 1000 rows
    table = sheet.tables.add(tableRange.getResizedRange(1000, 0), true);
    table.name = tableName;
    table.showTotals = false;
    table.getRange().format.wrapText = true;
    table.getRange().format.verticalAlignment = "Center";
    table.getHeaderRowRange().format.columnWidth = 220;
  }

  await context.sync();
}

async function addMultiSelectToLinkFieldsOnSheet(
  context: Excel.RequestContext,
  sheetName: ModelType | SFFModelType
) {
  // Get the sheet
  const sheet = context.workbook.worksheets.getItem(sheetName);

  // Get the total number of rows in the table
  // eslint-disable-next-line office-addins/load-object-before-read
  const totalRows = sheet.tables.getItem(sheetName).rows;
  totalRows.load("count");
  await context.sync();

  const newClass = createInstance(sheetName);
  const fields = newClass.getAllFields();

  for (const field of fields) {
    if (field.link) {
      await addMultiSelectToFields(
        context,
        sheet,
        field.displayName || field.name,
        field.link.table.className,
        totalRows.count
      );
    } else if (field.type === "select") {
      await addSelectToFields(
        context,
        sheet,
        field.displayName || field.name,
        [...predefinedCodeLists, ...hiddenSheets].find((listName) =>
          field.name.toLowerCase().includes(listName.toLowerCase())
        ) || field.name,
        totalRows.count
      );
    }
  }

  await context.sync();
}

async function addMultiSelectToFields(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  fieldName: string,
  linkTo: string,
  totalRows: number
) {
  const headerRange = sheet
    .getUsedRange()
    .find(fieldName, { completeMatch: true, matchCase: true });

  if (!headerRange) {
    if (dialogHandler) {
      dialogHandler(
        {
          descriptor: {
            id: "generics.error",
          },
        },
        {
          descriptor: {
            id: "createTables.messages.error.fieldNotFound",
            defaultMessage: "Field {fieldName} not found in the sheet {sheetName}.",
          },
          values: { fieldName, sheetName: sheet.name },
        }
      );
    }
    return;
  }

  // Get the column index of the field to add data validation
  headerRange.load("columnIndex");
  await context.sync();
  const column = headerRange.columnIndex;

  const range = sheet.getRangeByIndexes(1, column, totalRows, 1);

  // Add data validation to the range
  range.dataValidation.rule = {
    list: {
      inCellDropDown: true,
      source: `=INDIRECT("${linkTo}['@id]")`,
    },
  };

  // Block the user from entering values that are not in the list
  range.dataValidation.errorAlert = {
    message: "Please select a value from the list.",
    showAlert: true,
    style: "Stop",
    title: "Invalid value",
  };

  range.load("address");
  await context.sync();

  // Add multi-select change handler to the range
  addMultiSelectHandler(sheet, range.address);
}

async function addSelectToFields(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  fieldName: string,
  linkTo: string,
  totalRows: number
) {
  const headerRange = sheet
    .getUsedRange()
    .find(fieldName, { completeMatch: true, matchCase: true });

  if (!headerRange) {
    if (dialogHandler) {
      dialogHandler(
        {
          descriptor: {
            id: "generics.error",
          },
        },
        {
          descriptor: {
            id: "createTables.messages.error.fieldNotFound",
            defaultMessage: "Field {fieldName} not found in the sheet {sheetName}.",
          },
          values: { fieldName, sheetName: sheet.name },
        }
      );
    }
    return;
  }

  // Get the column index of the field to add data validation
  headerRange.load("columnIndex");
  await context.sync();
  const column = headerRange.columnIndex;

  const range = sheet.getRangeByIndexes(1, column, totalRows, 1);

  // Add data validation to the range
  range.dataValidation.rule = {
    list: {
      inCellDropDown: true,
      source: `=INDIRECT("${linkTo}[name]")`,
    },
  };

  // Block the user from entering values that are not in the list
  range.dataValidation.errorAlert = {
    message: "Please select a value from the list.",
    showAlert: true,
    style: "Stop",
    title: "Invalid value",
  };

  await context.sync();
}

function addMultiSelectHandler(sheet: Excel.Worksheet, rangeAddress: string) {
  sheet.onChanged.add(async (e) => multiSelectEventHandler(e, rangeAddress));
}

async function multiSelectEventHandler(
  event: Excel.WorksheetChangedEventArgs,
  rangeAddress: string
) {
  if (event.triggerSource === "ThisLocalAddin") {
    return;
  }
  if (!event.details) {
    if (dialogHandler) {
      dialogHandler(
        {
          descriptor: {
            id: "generics.warning",
          },
        },
        {
          descriptor: {
            id: "eventHandler.messages.warning.changesForRangeNotSupported",
            defaultMessage:
              "Changes for range are not supported. If you are trying to change fields that might be linked, please undo the operation and do it one by one.",
          },
        }
      );
    }
    return;
  }
  await Excel.run(async (context) => {
    try {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load("name");
      const targetCell = event.getRange(context);
      targetCell.load("address, values, formulas, rowIndex");
      await context.sync();
      const table = activeSheet.tables.getItem(activeSheet.name);
      const tableRange = table.getRange();
      const tableHeadersRange = table.getHeaderRowRange();
      tableHeadersRange.load("values");
      await context.sync();
      const idColumnIndex = tableHeadersRange.values[0].indexOf("@id");
      const idColumn = tableRange.getColumn(idColumnIndex);
      idColumn.load("values");
      const idRowIndex = targetCell.rowIndex;
      const idRow = tableRange.getRow(idRowIndex);
      const idCell = activeSheet.getRangeByIndexes(idRowIndex, idColumnIndex, 1, 1);
      idRow.load("values");
      idCell.load("values, address");
      await context.sync();
      const idColumnValues = idColumn.values;
      const idRowValues = idRow.values[0];

      // Check if change is in the id column
      if (idCell.address === targetCell.address) {
        const newId = event.details.valueAfter.toString();
        const oldId = event.details.valueBefore.toString();

        if (newId === "" || newId === null) {
          if (idRowValues.some((v: any) => v !== "")) {
            targetCell.values = [[oldId]];
            if (dialogHandler) {
              dialogHandler(
                {
                  descriptor: {
                    id: "generics.error",
                  },
                },
                {
                  descriptor: {
                    id: "eventHandler.messages.error.idCannotBeEmpty",
                    defaultMessage:
                      "The id cannot be empty for a row with values. Please remove all values first.",
                  },
                }
              );
            }
            return;
          }
          targetCell.values = [[newId]];
          return;
        }

        // Check if new id is unique
        if (
          idColumnValues.some((v: any[], i: number) => (idRowIndex === i ? false : v[0] === newId))
        ) {
          targetCell.values = [[oldId]];
          if (dialogHandler) {
            dialogHandler(
              {
                descriptor: {
                  id: "generics.error",
                },
              },
              {
                descriptor: {
                  id: "eventHandler.messages.error.newIdNotUnique",
                  defaultMessage: "The new id is not unique. Please enter a unique id.",
                },
              }
            );
          }
          return;
        }

        // Check if new id is a valid URL
        try {
          new URL(newId);
        } catch (error) {
          targetCell.values = [[oldId]];
          if (dialogHandler) {
            dialogHandler(
              {
                descriptor: {
                  id: "generics.error",
                },
              },
              {
                descriptor: {
                  id: "eventHandler.messages.error.newIdNotValidURL",
                  defaultMessage: "The new id is not a valid URL. Please enter a valid URL.",
                },
              }
            );
          }
          return;
        }

        targetCell.values = [[newId]];

        // Update related fields values
        const sheetName = activeSheet.name as ModelType | SFFModelType;
        const newClass = createInstance(sheetName);
        const fields = newClass.getAllFields();

        for (const field of fields) {
          if (
            field.link &&
            (!ignoredFields[field.link.table.className] ||
              !ignoredFields[field.link.table.className].includes(field.link.field))
          ) {
            await updateRelatedFieldsValues(
              context,
              field.link.table.className,
              field.link.field,
              oldId,
              newId
            );
          }
        }

        return;
      }

      if (isCellInRange(targetCell.address, rangeAddress)) {
        const newValue: string = event.details.valueAfter.toString();
        const oldValue: string = event.details.valueBefore.toString();
        const newValueArray = newValue.split(", ");
        const oldValueArray = oldValue.split(", ");

        // Get the field name or table header
        const columnLetter: string = targetCell.address.split("!")[1].split("")[0];
        const fieldRange = activeSheet.getRange(columnLetter + "1");
        fieldRange.load("values");
        await context.sync();
        const fieldName: string = fieldRange.values[0][0];
        const cid = createInstance(activeSheet.name as ModelType | SFFModelType);
        const field = cid.getFieldByName(fieldName);

        if (!idCell.values || idCell.values[0][0] === "") {
          targetCell.values = [[oldValue]];
          return;
        }

        if (!field || !field.link) {
          targetCell.values = [[oldValue]];
          if (dialogHandler) {
            dialogHandler(
              {
                descriptor: {
                  id: "generics.error",
                },
              },
              {
                descriptor: {
                  id: "eventHandler.messages.error.fieldNotFound",
                  defaultMessage: "Field not found.",
                },
              }
            );
          }
          return;
        }

        if (
          !(await checkIfAllValuesExistInRelatedSheet(
            context,
            field.link?.table.className,
            newValue
          ))
        ) {
          targetCell.values = [[oldValue]];
          if (dialogHandler) {
            dialogHandler(
              {
                descriptor: {
                  id: "generics.error",
                },
              },
              {
                descriptor: {
                  id: "eventHandler.messages.error.invalidValue",
                  defaultMessage: "Invalid value. Please select a value from the list.",
                },
              }
            );
          }
          return;
        }

        // Avoid second event trigger
        if (
          !(await checkIfAllValuesExistInRelatedSheet(
            context,
            field.link?.table.className,
            oldValue
          ))
        ) {
          return;
        }

        let targetCellValue;

        if ((newValue === "" || newValue === null) && (oldValue === "" || oldValue === null)) {
          targetCellValue = [[newValue.toString()]];
        }
        if (newValueArray.length > 1 && oldValueArray.length > 1 && newValue === oldValue) {
          targetCellValue = [[oldValue]];
          return;
        }
        if (oldValue && oldValueArray.indexOf(newValue) === -1) {
          targetCellValue = [[oldValue + ", " + newValue]];
        } else if (oldValue && oldValueArray.indexOf(newValue) !== -1) {
          const newValues = oldValueArray.filter((value: string) => value !== newValue);
          targetCellValue = [[newValues.join(", ")]];
        }
        if (!oldValue) {
          targetCellValue = [[newValue]];
        }

        targetCell.values = targetCellValue || [[oldValue]];

        if (
          field.link.table.className !== activeSheet.name &&
          (!ignoredFields[field.link.table.className] ||
            !ignoredFields[field.link.table.className].includes(field.link.field))
        ) {
          await handleLinkedFieldsInRelatedSheet(
            context,
            field.link?.table.className,
            field.link?.field,
            idCell.values[0][0].toString(),
            targetCellValue ? targetCellValue[0][0] : oldValue.toString()
          );
        }
      }
    } catch (error) {
      console.log("Error: " + error);
    }
  });
}

// Check if target cell is in the multi-select range
function isCellInRange(cellAddress: string, columnsRangeAddress: string) {
  const [startCell, endCell] = columnsRangeAddress.split(":");
  const columnRangeStart = startCell.match(/[A-Z]\d+/g) || [];
  const columnRangeEnd = endCell.match(/[A-Z]\d+/g) || [];
  if (
    columnRangeStart.length <= 0 ||
    columnRangeEnd.length <= 0 ||
    !columnRangeStart[0] ||
    !columnRangeEnd[0]
  ) {
    return false;
  }
  const startColumn = columnRangeStart[0].slice(0, 1);
  const endColumn = columnRangeEnd[0].slice(0, 1);
  const startRow = columnRangeStart[0].slice(1);
  const endRow = columnRangeEnd[0].slice(1);

  const cell = cellAddress.split("!")[1];

  const cellColumn = cell.slice(0, 1);
  const cellRow = cell.slice(1);

  if (!startColumn || !startRow || !endColumn || !endRow || !cellColumn || !cellRow) {
    return false;
  }
  return (
    cellColumn >= startColumn &&
    cellColumn <= endColumn &&
    +cellRow >= +startRow &&
    +cellRow <= +endRow
  );
}

async function addMultiSelectHandlerToAllTables() {
  await Excel.run(async (context) => {
    try {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");

      await context.sync();

      const sheetNames = Object.keys(map);

      for (const sheet of sheets.items) {
        if (sheetNames.includes(sheet.name)) {
          // Add multi-select handler to the link fields on tables
          const newClass = createInstance(sheet.name as ModelType | SFFModelType);
          const fields = newClass.getAllFields();

          // Get the total number of rows in the table
          // eslint-disable-next-line office-addins/load-object-before-read
          const totalRows = sheet.tables.getItem(sheet.name).rows;
          totalRows.load("count");
          await context.sync();

          // Add multi-select functionality to link fields
          for (const field of fields) {
            if (field.link) {
              const headerRange = sheet
                .getUsedRange()
                .find(field.displayName || field.name, { completeMatch: true, matchCase: true });

              if (!headerRange) {
                if (dialogHandler) {
                  dialogHandler(
                    {
                      descriptor: {
                        id: "generics.error",
                      },
                    },
                    {
                      descriptor: {
                        id: "createTables.messages.error.fieldNotFound",
                        defaultMessage: "Field {fieldName} not found in the sheet {sheetName}.",
                      },
                      values: { fieldName: field.name, sheetName: sheet.name },
                    }
                  );
                }
                return;
              }
              // Get the column index of the field to add data validation
              headerRange.load("columnIndex");
              await context.sync();
              const column = headerRange.columnIndex;

              const range = sheet.getRangeByIndexes(1, column, totalRows.count, 1);

              range.load("address");
              await context.sync();
              addMultiSelectHandler(sheet, range.address);
            }
          }
        }
      }

      await context.sync();
    } catch (error) {
      console.log("Error: " + error);
    }
  });
}

export async function populateSelectLists() {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;

    for (const codeList of ["StreetType", "StreetDirection"]) {
      try {
        const sheet = sheets.getItem(codeList);
        const table = sheet.tables.getItem(codeList);
        const range = table.getRange();
        range.load("address");
        await context.sync();

        await populateCodeList(context, codeList, range.address, codeList);
      } catch (error) {
        console.log("Error: " + error);
        throw new Error(`${codeList}`);
      }
    }

    await context.sync();
  });
}

export async function populateCodeLists() {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;

    for (const codeList of predefinedCodeLists) {
      try {
        const sheet = sheets.getItem(codeList);
        const table = sheet.tables.getItem(codeList);
        const range = table.getRange();
        range.load("address");
        await context.sync();

        await populateCodeList(context, codeList, range.address, codeList);
      } catch (error) {
        console.log("Error: " + error);
        throw new Error(`${codeList}`);
      }
    }

    await context.sync();
  });
}

async function populateCodeList(
  context: Excel.RequestContext,
  tableName: string,
  rangeAddress: string,
  codeList: string
) {
  const range = context.workbook.worksheets.getItem(tableName).getRange(rangeAddress);
  range.load("values");
  await context.sync();

  const data = await getCodeListByTableName(codeList);

  const headers = hiddenSheets.includes(tableName) ? ["id", "name"] : Object.keys(data[0]);

  // Check if the data is already populated if so, update the data
  for (const row of data) {
    const idColumnIndex = range.values[0].indexOf(hiddenSheets.includes(tableName) ? "id" : "@id");
    const idColumn = range.getColumn(idColumnIndex);
    idColumn.load("values");
    await context.sync();
    const idColumnValues = idColumn.values.map((v: any) => v[0]);

    let rowIndex = idColumnValues.indexOf(row["@id"]);
    if (rowIndex === -1) {
      rowIndex = idColumnValues.indexOf("");
    }

    const rowRange = range.getRow(rowIndex);
    range.load("values");
    await context.sync();
    for (const header of headers) {
      const headerIndex = range.values[0].indexOf(header);
      let propHeader = header;
      if (header === "id") {
        propHeader = "@id";
      } else if (header === "name") {
        propHeader = "hasName";
      }
      rowRange.getCell(0, headerIndex).values = [[(row as any)[propHeader]]];
    }
    await context.sync();
  }

  await context.sync();
}
