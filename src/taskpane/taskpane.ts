import { dialogHandler } from './context/DialogContext';
import { createInstance, ignoredFields, map, ModelType } from './domain/models';
import {
  checkIfAllValuesExistInRelatedSheet,
  handleLinkedFieldsInRelatedSheet,
  updateRelatedFieldsValues,
} from './helpers/handleLinkedFieldsOnOtherSheet';

/* global Office */
Office.onReady(() => {
  // If needed, Office.js is ready to be called.

  // If the user is in Excel, add multi-select functionality if the user have all the standard tables
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
          console.log('Error: ' + error);
        }
      }
      await context.sync();

      // Add tables to the sheets
      for (const sheetName of Object.keys(map)) {
        try {
          await createTable(context, sheetName as ModelType);
        } catch (error) {
          console.log('Error: ' + error);
        }
      }

      // Add multi-select functionality to link fields
      for (const sheetName of Object.keys(map)) {
        try {
          await addMultiSelectToLinkFieldsOnSheet(context, sheetName as ModelType);
        } catch (error) {
          console.log('Error: ' + error);
        }
      }

      await context.sync();
    } catch (error) {
      console.log('Error: ' + error);
    }
  });
}

async function createTable(context: Excel.RequestContext, sheetName: ModelType) {
  const sheet = context.workbook.worksheets.getItem(sheetName);

  const tableName = sheetName.toString();
  const newClass = createInstance(sheetName);
  const fields = newClass.getFields();

  // If header is '@id', replace it with "'@id" to avoid Excel error
  const headers = fields.map((field) => (field.name === '@id' ? "'@id" : field.name)); // eslint-disable-line quotes

  // Append helper fields for multi-select
  if (ignoredFields[tableName]) {
    for (const field of ignoredFields[tableName]) {
      headers.push(field);
    }
  }

  // Add a table to the sheet
  const rangeAddress = `A1:${String.fromCharCode(65 + headers.length - 1)}1`;
  const tableRange = sheet.getRange(rangeAddress);
  tableRange.values = [headers];
  tableRange.format.autofitColumns();

  // Resize the range to 1000 rows
  const table = sheet.tables.add(tableRange.getResizedRange(1000, 0), true);
  table.name = tableName;
  table.showTotals = false;

  await context.sync();
}

async function addMultiSelectToLinkFieldsOnSheet(
  context: Excel.RequestContext,
  sheetName: ModelType
) {
  // Get the sheet
  const sheet = context.workbook.worksheets.getItem(sheetName);

  // Get the total number of rows in the table
  // eslint-disable-next-line office-addins/load-object-before-read
  const totalRows = sheet.tables.getItem(sheetName).rows;
  totalRows.load('count');
  await context.sync();

  const newClass = createInstance(sheetName);
  const fields = newClass.getFields();

  for (const field of fields) {
    if (field.link) {
      await addMultiSelect(context, sheet, field.name, field.link.className, totalRows.count);
    }
  }

  // Add multi-select to the helper fields
  if (ignoredFields[sheetName]) {
    for (const field of ignoredFields[sheetName]) {
      await addMultiSelect(context, sheet, field, field.substring(3), totalRows.count);
    }
  }

  await context.sync();
}

async function addMultiSelect(
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
      dialogHandler('Error', `Field ${fieldName} not found in the sheet.`);
    }
    return;
  }

  // Get the column index of the field to add data validation
  headerRange.load('columnIndex');
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
    message: 'Please select a value from the list.',
    showAlert: true,
    style: 'Stop',
    title: 'Invalid value',
  };

  range.load('address');
  await context.sync();

  // Add multi-select change handler to the range
  addMultiSelectHandler(sheet, range.address);
}

function addMultiSelectHandler(sheet: Excel.Worksheet, rangeAddress: string) {
  sheet.onChanged.add(async (e) => multiSelectEventHandler(e, rangeAddress));
}

async function multiSelectEventHandler(
  event: Excel.WorksheetChangedEventArgs,
  rangeAddress: string
) {
  if (event.triggerSource === 'ThisLocalAddin') {
    return;
  }
  await Excel.run(async (context) => {
    try {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load('name');
      const targetCell = event.getRange(context);
      targetCell.load('address, values, formulas, rowIndex');
      await context.sync();
      const table = activeSheet.tables.getItem(activeSheet.name);
      const tableRange = table.getRange();
      const tableHeadersRange = table.getHeaderRowRange();
      tableHeadersRange.load('values');
      await context.sync();
      const idColumnIndex = tableHeadersRange.values[0].indexOf('@id');
      const idColumn = tableRange.getColumn(idColumnIndex);
      idColumn.load('values');
      const idRowIndex = targetCell.rowIndex;
      const idCell = activeSheet.getRangeByIndexes(idRowIndex, idColumnIndex, 1, 1);
      idCell.load('values, address');
      await context.sync();
      const idColumnValues = idColumn.values;

      // Check if change is in the id column
      if (idCell.address === targetCell.address) {
        const newId = event.details.valueAfter.toString();
        const oldId = event.details.valueBefore.toString();

        // Check if new id is unique
        if (newId === '' || newId === null) {
          targetCell.values = [[oldId]];
          if (dialogHandler) {
            dialogHandler('Error', 'The new id cannot be empty. Please enter a valid id.');
          }
          return;
        }

        if (
          idColumnValues.some((v: any[], i: number) => (idRowIndex === i ? false : v[0] === newId))
        ) {
          targetCell.values = [[oldId]];
          if (dialogHandler) {
            dialogHandler('Error', 'The new id is not unique. Please enter a unique id.');
          }
          return;
        }

        // Check if new id is a valid URL
        try {
          new URL(newId);
        } catch (error) {
          targetCell.values = [[oldId]];
          if (dialogHandler) {
            dialogHandler('Error', 'The new id is not a valid URL. Please enter a valid URL.');
          }
          return;
        }

        targetCell.values = [[newId]];

        // Update related fields values
        const sheetName = activeSheet.name;
        const newClass = createInstance(sheetName as ModelType);
        const fields = newClass.getFields();

        for (const field of fields) {
          if (field.link) {
            const relatedClass = createInstance(field.link.className as ModelType);
            const relatedFields = relatedClass.getFields();
            const relatedFieldName = relatedFields.find((f) => f.link?.className === sheetName);
            if (relatedFieldName) {
              await updateRelatedFieldsValues(
                context,
                field.link.className,
                relatedFieldName.name,
                oldId,
                newId
              );
            }
          }
        }

        // Update related helper fields values
        if (ignoredFields[sheetName]) {
          for (const field of ignoredFields[sheetName]) {
            const fieldName = field.startsWith('has') ? 'for' + sheetName : 'has' + sheetName;
            await updateRelatedFieldsValues(context, field.substring(3), fieldName, oldId, newId);
          }
        }

        return;
      }

      if (isCellInRange(targetCell.address, rangeAddress)) {
        const newValue: string = event.details.valueAfter.toString();
        const oldValue: string = event.details.valueBefore.toString();
        const newValueArray = newValue.split(', ');
        const oldValueArray = oldValue.split(', ');

        // Get the field name or table header
        const columnLetter: string = targetCell.address.split('!')[1].split('')[0];
        const fieldRange = activeSheet.getRange(columnLetter + '1');
        fieldRange.load('values');
        await context.sync();
        const fieldName: string = fieldRange.values[0][0];

        if (!idCell.values || idCell.values[0][0] === '') {
          targetCell.values = [[oldValue]];
          return;
        }

        if (!(await checkIfAllValuesExistInRelatedSheet(context, fieldName.slice(3), newValue))) {
          targetCell.values = [[oldValue]];
          if (dialogHandler) {
            dialogHandler('Error', 'Invalid value. Please select a value from the list.');
          }
          return;
        }

        // Avoid second event trigger
        if (!(await checkIfAllValuesExistInRelatedSheet(context, fieldName.slice(3), oldValue))) {
          return;
        }

        let targetCellValue;

        if ((newValue === '' || newValue === null) && (oldValue === '' || oldValue === null)) {
          targetCellValue = [[newValue.toString()]];
        }
        if (newValueArray.length > 1 && oldValueArray.length > 1 && newValue === oldValue) {
          targetCellValue = [[oldValue]];
          return;
        }
        if (oldValue && oldValueArray.indexOf(newValue) === -1) {
          targetCellValue = [[oldValue + ', ' + newValue]];
        } else if (oldValue && oldValueArray.indexOf(newValue) !== -1) {
          const newValues = oldValueArray.filter((value: string) => value !== newValue);
          targetCellValue = [[newValues.join(', ')]];
        }
        if (!oldValue) {
          targetCellValue = [[newValue]];
        }

        targetCell.values = targetCellValue || [[oldValue]];

        await handleLinkedFieldsInRelatedSheet(
          context,
          fieldName.slice(3),
          `${fieldName.startsWith('has') ? 'for' : 'has'}${activeSheet.name}`,
          idCell.values[0][0].toString(),
          targetCellValue ? targetCellValue[0][0] : oldValue.toString()
        );
      }
    } catch (error) {
      console.log('Error: ' + error);
    }
  });
}

// Check if target cell is in the multi-select range
function isCellInRange(cellAddress: string, columnsRangeAddress: string) {
  const [startCell, endCell] = columnsRangeAddress.split(':');
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

  const cell = cellAddress.split('!')[1];

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
      sheets.load('items/name');

      await context.sync();

      const sheetNames = Object.keys(map);

      for (const sheet of sheets.items) {
        if (sheetNames.includes(sheet.name)) {
          // Add multi-select handler to the link fields on tables
          const newClass = createInstance(sheet.name as ModelType);
          const fields = newClass.getFields();

          // Get the total number of rows in the table
          // eslint-disable-next-line office-addins/load-object-before-read
          const totalRows = sheet.tables.getItem(sheet.name as ModelType).rows;
          totalRows.load('count');
          await context.sync();

          // Add multi-select functionality to link fields
          for (const field of fields) {
            if (field.link) {
              const headerRange = sheet
                .getUsedRange()
                .find(field.name, { completeMatch: true, matchCase: true });

              if (!headerRange) {
                if (dialogHandler) {
                  dialogHandler('Error', `Field ${field.name} not found in the sheet.`);
                }
                return;
              }
              // Get the column index of the field to add data validation
              headerRange.load('columnIndex');
              await context.sync();
              const column = headerRange.columnIndex;

              const range = sheet.getRangeByIndexes(1, column, totalRows.count, 1);

              range.load('address');
              await context.sync();
              addMultiSelectHandler(sheet, range.address);
            }
          }
        }
      }

      await context.sync();
    } catch (error) {
      console.log('Error: ' + error);
    }
  });
}
