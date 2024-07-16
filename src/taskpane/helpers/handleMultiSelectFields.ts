import { mapToHiddenLinkSheet } from '../domain/models';

/* global Excel console */
export async function handleMultiSelectFields(
  event: Excel.WorksheetChangedEventArgs,
  context: Excel.RequestContext
) {
  const targetTable = context.workbook.worksheets.getItem(event.worksheetId);
  const targetRange = event.getRange(context);
  targetTable.load('name');
  targetRange.load('values');
  await context.sync();
  const tableName = targetTable.name;

  // Get related tables that have multi-select fields
  const relatedTables = [];
  for (const sheetName of Object.keys(mapToHiddenLinkSheet)) {
    for (const fieldName of Object.keys(mapToHiddenLinkSheet[sheetName])) {
      if (mapToHiddenLinkSheet[sheetName][fieldName] === tableName) {
        relatedTables.push({ sheetName, fieldName });
      }
    }
  }

  if (relatedTables.length === 0) {
    return;
  }

  for (const { sheetName, fieldName } of relatedTables) {
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    worksheet.load('tables');
    await context.sync();
    const table = worksheet.tables.getItem(sheetName);

    const tableRange = table.getRange();
    const tableIdColumn = tableRange.getColumn(0);
    const tableHeaders = table.getHeaderRowRange();
    tableHeaders.load('values');
    tableIdColumn.load('values');
    await context.sync();
    const fieldNameIndex = tableHeaders.values[0].indexOf(fieldName);
    const idColumnValues = tableIdColumn.values;

    let targetCell: Excel.Range;
    let newValue = '';
    let oldValue = '';

    // if field starts with 'has' for each id in the field concatenate the values found in the hidden link sheet
    if (fieldName.startsWith('has')) {
      const id = targetRange.values[0][0];
      console.log('id', id);
      const idRowIndex = idColumnValues.findIndex((row) => row[0] === id);
      targetCell = tableRange.getCell(idRowIndex, fieldNameIndex);
      targetCell.load('values');
      await context.sync();
      oldValue = targetCell.values[0][0];
      newValue = targetRange.values[0][1];

      console.log('oldValue', oldValue);
      console.log('newValue', newValue);
    } else {
      // if field starts with 'for' for each id in the field concatenate the values found in the hidden link sheet
      const id = targetRange.values[0][1];
      console.log('id', id);
      const idRowIndex = idColumnValues.findIndex((row) => row[0] === id);
      targetCell = tableRange.getCell(idRowIndex, fieldNameIndex);
      targetCell.load('values');
      await context.sync();
      oldValue = targetCell.values[0][0];
      newValue = targetRange.values[0][0];

      console.log('oldValue', oldValue);
      console.log('newValue', newValue);
    }

    if ((newValue === '' || newValue === null) && (oldValue === '' || oldValue === null)) {
      continue;
    }
    if (oldValue && oldValue.toString().indexOf(newValue.toString()) === -1) {
      targetCell.values = [[oldValue.toString() + ', ' + newValue.toString()]];
    }
    // } else if (oldValue && oldValue.toString().indexOf(newValue.toString()) !== -1) {
    //   targetCell.values = [
    //     [
    //       oldValue
    //         .replace(newValue, '')
    //         .replace(/^, |^,/g, '')
    //         .replace(', ,', '')
    //         .trim(),
    //     ],
    //   ];
    // }
    if (!oldValue) {
      targetCell.values = [[newValue.toString()]];
    }
  }
  await context.sync();
}
