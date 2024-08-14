/* global Excel console */
export async function handleLinkedFieldsInRelatedSheet(
  context: Excel.RequestContext,
  relatedTableName: string,
  relatedFieldName: string,
  id: string,
  value: string
) {
  try {
    // Find the related field column in the related table
    const worksheet = context.workbook.worksheets.getItem(relatedTableName);
    worksheet.load('tables');
    await context.sync();
    const table = worksheet.tables.getItem(relatedTableName);
    const tableHeadersRange = table.getHeaderRowRange();
    const tableRange = table.getRange();
    tableHeadersRange.load('values');
    tableRange.load('values');
    await context.sync();

    // Get teh index of the related field in the related table
    const relatedFieldIndex = tableHeadersRange.values[0].indexOf(relatedFieldName);
    const idColumnIndex = tableHeadersRange.values[0].indexOf('@id');

    // For all cells in the related field column, check if the cell value is stringified array and if the id is in the array
    const relatedFieldColumn = tableRange.getColumn(relatedFieldIndex);
    const idColumn = tableRange.getColumn(idColumnIndex);
    relatedFieldColumn.load('values');
    idColumn.load('values');
    await context.sync();
    const relatedFieldValues = relatedFieldColumn.values;
    const idColumnValues = idColumn.values;

    // Break value into array
    const valueArray = value ? value.split(', ') : [];

    // For each cell in the related field column, check if the id is in the array
    for (let i = 0; i < relatedFieldValues.length; i++) {
      const idColumnValue = idColumnValues[i][0].toString();
      const relatedFieldValue = relatedFieldValues[i][0].toString();
      const relatedFieldValueArray: string[] = relatedFieldValue.split(', ');

      if (!idColumnValue) {
        continue;
      }

      if (!value && relatedFieldValue && relatedFieldValueArray.includes(id)) {
        // Remove the id from the array
        const newValueArray = relatedFieldValueArray.filter((v: string) => v !== id);
        const newValue = newValueArray.join(', ');
        relatedFieldColumn.getCell(i, 0).values = [[newValue]];
      } else if (
        value &&
        !valueArray.includes(idColumnValue) &&
        relatedFieldValue &&
        relatedFieldValueArray.includes(id)
      ) {
        // Remove the id from the array
        const newValueArray = relatedFieldValueArray.filter((v: string) => v !== id);
        const newValue = newValueArray.join(', ');
        relatedFieldColumn.getCell(i, 0).values = [[newValue]];
      } else if (
        value &&
        valueArray.includes(idColumnValue) &&
        (!relatedFieldValue || !relatedFieldValueArray.includes(id))
      ) {
        // Add the id to the array
        const newValueArray = relatedFieldValue ? relatedFieldValueArray : [];
        newValueArray.push(id);
        const newValue = newValueArray.join(', ');
        relatedFieldColumn.getCell(i, 0).values = [[newValue]];
      }
    }

    // Save the changes
    await context.sync();
  } catch (error) {
    console.error(error);
  }
}

// Handler to check if all values in the new values from the event exist in the relate sheet @id column
export async function checkIfAllValuesExistInRelatedSheet(
  context: Excel.RequestContext,
  relatedTableName: string,
  values: string
) {
  try {
    const worksheet = context.workbook.worksheets.getItem(relatedTableName);
    worksheet.load('tables');
    await context.sync();
    const table = worksheet.tables.getItem(relatedTableName);
    const tableHeadersRange = table.getHeaderRowRange();
    const tableRange = table.getRange();
    tableHeadersRange.load('values');
    tableRange.load('values');
    await context.sync();

    // Get the index of the @id field in the related table
    const idColumnIndex = tableHeadersRange.values[0].indexOf('@id');

    // For all values check if they exist in the related table
    const idColumn = tableRange.getColumn(idColumnIndex);
    idColumn.load('values');
    await context.sync();
    const idColumnValues = idColumn.values;

    const valuesArray = values.split(', ');

    for (let i = 0; i < valuesArray.length; i++) {
      const value = valuesArray[i];
      const valueExists = idColumnValues.some((v: any[]) => v[0] === value);
      if (!valueExists) {
        return false;
      }
    }

    return true;
  } catch (error) {
    console.error(error);
    return false;
  }
}

// Handler to update related fields values when an id is changed
export async function updateRelatedFieldsValues(
  context: Excel.RequestContext,
  relatedTableName: string,
  relatedFieldName: string,
  oldValue: string,
  newValue: string
) {
  try {
    // Find the related field column in the related table
    const worksheet = context.workbook.worksheets.getItem(relatedTableName);
    worksheet.load('tables');
    await context.sync();
    const table = worksheet.tables.getItem(relatedTableName);
    const tableHeadersRange = table.getHeaderRowRange();
    const tableRange = table.getRange();
    tableHeadersRange.load('values');
    tableRange.load('values');
    await context.sync();

    // Get the index of the related field in the related table
    const relatedFieldIndex = tableHeadersRange.values[0].indexOf(relatedFieldName);

    // For all cells in the related field column, check if the cell value is stringified array and if the id is in the array
    const relatedFieldColumn = tableRange.getColumn(relatedFieldIndex);
    relatedFieldColumn.load('values');
    await context.sync();
    const relatedFieldValues = relatedFieldColumn.values;

    // For each cell in the related field column, check if the id is in the array
    for (let i = 0; i < relatedFieldValues.length; i++) {
      const relatedFieldValue = relatedFieldValues[i][0].toString();
      const relatedFieldValueArray: string[] = relatedFieldValue.split(', ');

      if (!relatedFieldValue) {
        continue;
      }

      if (relatedFieldValueArray.includes(oldValue)) {
        // update the id in the array
        const newValueArray = relatedFieldValueArray.map((v: string) =>
          v === oldValue ? newValue : v
        );
        const newValues = newValueArray.join(', ');
        relatedFieldColumn.getCell(i, 0).values = [[newValues]];
      }
    }

    // Save the changes
    await context.sync();
  } catch (error) {
    console.error(error);
  }
}
