import { mapToHiddenLinkSheet } from '../domain/models';

/* global Excel console */
export async function handleLinkedFields(
  event: Excel.WorksheetChangedEventArgs,
  context: Excel.RequestContext
) {
  // Get the active sheet and target cell
  const activeSheet = context.workbook.worksheets.getActiveWorksheet();
  const targetCell = activeSheet.getRange(event.address);
  const idCell = activeSheet.getRange('A' + event.address.slice(1));
  activeSheet.load('name, values');
  targetCell.load('address, values, formulas');
  idCell.load('values');
  await context.sync();

  // Filed name or table header
  const columnLetter: string = targetCell.address.split('!')[1].split('')[0];
  const fieldRange = activeSheet.getRange(columnLetter + '1');
  fieldRange.load('values');
  await context.sync();
  const fieldName: string = fieldRange.values[0][0];
  const tableName = activeSheet.name;
  const id = idCell.values[0][0];

  // Get hidden link sheet
  const hiddenLinkSheet = context.workbook.worksheets.getItem(
    mapToHiddenLinkSheet[tableName][fieldName]
  );
  hiddenLinkSheet.load('name');
  await context.sync();

  // unprotected the hidden link sheet
  hiddenLinkSheet.protection.unprotect();
  hiddenLinkSheet.visibility = 'Visible';
  await context.sync();

  // Get trigger values
  const valueBefore = event.details.valueBefore.toString();
  const valueAfter = event.details.valueAfter.toString();

  // Check if triggered id is already in the hidden link sheet
  let isTargetValueInHiddenLinkSheet: Excel.RangeAreas | null = hiddenLinkSheet.findAllOrNullObject(
    valueAfter || valueBefore,
    {
      completeMatch: true,
      matchCase: true,
    }
  );
  await context.sync();
  if (isTargetValueInHiddenLinkSheet.isNullObject) {
    isTargetValueInHiddenLinkSheet = null;
  }

  // If id is already related to the target value, remove the relation
  let valueFound = false;
  if (isTargetValueInHiddenLinkSheet) {
    isTargetValueInHiddenLinkSheet.load('address, areas');
    await context.sync();

    console.log('isTargetValueInHiddenLinkSheet:', isTargetValueInHiddenLinkSheet);

    // For each address check if the id is related to the target value
    // if field name start with "has" search column A if starts with "for" search column B
    for (const area of isTargetValueInHiddenLinkSheet.areas.items) {
      if (fieldName.startsWith('has')) {
        if (/^A/.test(area.address.split('!')[1] as string)) {
          continue;
        }

        const cell = area.address.split('!')[1];
        const rowIndex = cell.split('')[1];
        const idRange = hiddenLinkSheet.getRange('A' + rowIndex);
        idRange.load('values');
        await context.sync();
        if (idRange.values[0][0] === id) {
          valueFound = true;
          await cleanRelationValues(hiddenLinkSheet, rowIndex);
        }
      } else if (fieldName.startsWith('for')) {
        if (/^B/.test(area.address.split('!')[1] as string)) {
          continue;
        }

        const cell = area.address.split('!')[1];
        const rowIndex = cell.split('')[1];
        const idRange = hiddenLinkSheet.getRange('B' + rowIndex);
        idRange.load('values');
        await context.sync();
        if (idRange.values[0][0] === id) {
          valueFound = true;
          await cleanRelationValues(hiddenLinkSheet, rowIndex);
        }
      }
    }

    // Delete the empty rows
    await deleteEmptyRows(hiddenLinkSheet, context);
    await context.sync();
  }

  if (!valueFound && isTargetValueInHiddenLinkSheet) {
    // If id is not related to the target value, add the relation
    await addRelation(hiddenLinkSheet, id, event, fieldName, context);
  }

  if (valueAfter && !isTargetValueInHiddenLinkSheet) {
    // If id is not related to the target value, add the relation
    await addRelation(hiddenLinkSheet, id, event, fieldName, context);
  }

  // Protect the hidden link sheet
  hiddenLinkSheet.protection.protect();
  hiddenLinkSheet.visibility = 'Hidden';
  await context.sync();
}

async function cleanRelationValues(hiddenLinkSheet: Excel.Worksheet, rowIndex: string) {
  const range = hiddenLinkSheet.getRange('A' + rowIndex + ':B' + rowIndex);
  range.values = [['', '']];
}

async function addRelation(
  hiddenLinkSheet: Excel.Worksheet,
  id: string,
  event: Excel.WorksheetChangedEventArgs,
  fieldName: string,
  context: Excel.RequestContext
) {
  const lastRow = hiddenLinkSheet.getUsedRange().getLastRow();
  lastRow.load('rowIndex');
  await context.sync();
  let nextRow = lastRow.rowIndex + 1;

  // check if last row is empty
  let range = hiddenLinkSheet.getRange('A' + nextRow + ':B' + nextRow);
  range.load('values');
  await context.sync();

  // if row is not empty, get the next empty row
  if (range.values[0][0] && range.values[0][1]) {
    nextRow += 1;
    range = hiddenLinkSheet.getRange('A' + nextRow + ':B' + nextRow);
  }

  range.load('address');
  await context.sync();

  if (fieldName.startsWith('has')) {
    range.values = [[id, event.details.valueAfter]];
  } else if (fieldName.startsWith('for')) {
    range.values = [[event.details.valueAfter, id]];
  }
  await context.sync();
}

async function deleteEmptyRows(hiddenLinkSheet: Excel.Worksheet, context: Excel.RequestContext) {
  const usedRange = hiddenLinkSheet.getUsedRange();
  usedRange.load('rowCount');
  await context.sync();
  const rowCount = usedRange.rowCount;

  for (let i = rowCount; i > 0; i--) {
    const range = hiddenLinkSheet.getRange('A' + i + ':B' + i);
    range.load('values');
    await context.sync();

    if (!range.values[0][0] && !range.values[0][1]) {
      range.delete(Excel.DeleteShiftDirection.up);
    }
  }
}
