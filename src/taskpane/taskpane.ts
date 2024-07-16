import { createInstance, hiddenLinkSheets, map, ModelType } from './domain/models';
import { handleLinkedFields } from './helpers/handleLikedFields';
import { handleMultiSelectFields } from './helpers/handleMultiSelectFields';

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

      // Create hidden link sheets
      for (const sheetName of hiddenLinkSheets) {
        try {
          await createHiddenLinkSheet(context, sheetName);
        } catch (error) {
          console.log('Error: ' + error);
        }
      }

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

      // Add event handler to hidden link sheets
      for (const sheetName of hiddenLinkSheets) {
        try {
          const sheet = context.workbook.worksheets.getItem(sheetName);
          addHiddenLinkSheetHandler(sheet);
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

  // If hearer is '@id', replace it with "'@id" to avoid Excel error
  const headers = fields.map((field) => (field.name === '@id' ? "'@id" : field.name)); // eslint-disable-line quotes

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

async function createHiddenLinkSheet(context: Excel.RequestContext, sheetName: string) {
  const sheets = context.workbook.worksheets;
  sheets.load('items/name');

  await context.sync();

  if (sheets.items.find((sheet) => sheet.name === sheetName)) {
    return;
  }

  const sheet = sheets.add(sheetName);
  sheet.visibility = 'Hidden';
  sheet.protection.protect();

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
    console.log('Field not found in the sheet.');
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
  addMultiSelectHandler(sheet, range.address);
}

function addMultiSelectHandler(sheet: Excel.Worksheet, rangeAddress: string) {
  sheet.onChanged.add(async (event) => {
    if (event.triggerSource === 'ThisLocalAddin') {
      return;
    }
    // console.log('event.changeType: ' + event.changeType);
    // console.log('event: ' + event.address);
    // console.log('event: ' + JSON.stringify(event.details));
    // console.log('event: ' + JSON.stringify(event));
    // To do: handle event.changeType RangeEdited for deleted or changed ids in @id field
    // To do: handle event.changeType RowDeleted for deleted rows
    await Excel.run(async (context) => {
      try {
        const activeSheet = context.workbook.worksheets.getActiveWorksheet();
        const targetCell = event.getRange(context);
        const idCell = activeSheet.getRange('A' + event.address.slice(1));
        targetCell.load('address, values, formulas');
        idCell.load('values');
        await context.sync();

        if (isCellInRange(targetCell.address, rangeAddress)) {
          const newValue = event.details.valueAfter.toString();
          const oldValue = event.details.valueBefore.toString();

          if (!idCell.values || idCell.values[0][0] === '') {
            targetCell.values = [[event.details.valueBefore.toString()]];
            return;
          }
          if ((newValue === '' || newValue === null) && (oldValue === '' || oldValue === null)) {
            targetCell.values = [[newValue.toString()]];
          }
          if (oldValue && oldValue.toString().indexOf(newValue.toString()) === -1) {
            targetCell.values = [[oldValue.toString() + ', ' + newValue.toString()]];
          } else if (oldValue && oldValue.toString().indexOf(newValue.toString()) !== -1) {
            targetCell.values = [
              [
                oldValue
                  .replace(newValue, '')
                  .replace(/^, |^,/g, '')
                  .replace(', ,', '')
                  .trim(),
              ],
            ];
          }
          if (!oldValue) {
            targetCell.values = [[newValue.toString()]];
          }
          await handleLinkedFields(event, context);
        }
      } catch (error) {
        console.log('Error: ' + error);
      }
    });
  });
}

function addHiddenLinkSheetHandler(sheet: Excel.Worksheet) {
  sheet.onChanged.add(async (event) => {
    await Excel.run(async (context) => {
      try {
        console.log('Event triggered:', event.address);
        console.log('Event: ', JSON.stringify(event));
        await handleMultiSelectFields(event, context);
      } catch (error) {
        console.log('Error: ' + error);
      }
    });
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
                console.log('Field not found in the sheet.');
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

      for (const sheet of sheets.items) {
        if (hiddenLinkSheets.includes(sheet.name)) {
          addHiddenLinkSheetHandler(sheet);
        }
      }

      await context.sync();
    } catch (error) {
      console.log('Error: ' + error);
    }
  });
}
