/* global Excel console */

/**
 * Updates bidirectional relationships between tables when values change
 */
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

    // Check if the worksheet/table exists before proceeding
    try {
      const table = worksheet.tables.getItem(relatedTableName);
      const tableHeadersRange = table.getHeaderRowRange();
      tableHeadersRange.load("values");
      await context.sync();

      // Get field indices
      const relatedFieldIndex = tableHeadersRange.values[0].indexOf(relatedFieldName);
      const idColumnIndex = tableHeadersRange.values[0].indexOf("@id");

      if (relatedFieldIndex === -1) {
        return;
      }

      if (idColumnIndex === -1) {
        return;
      }

      // Use DataBodyRange for more efficient access to table data
      const dataRange = table.getDataBodyRange();
      // Fix: Explicitly load all the data we need in one batch
      dataRange.load("values");
      dataRange.load("rowCount");
      await context.sync();

      // Store values in a local variable to avoid repeated property access
      const dataValues = dataRange.values;
      const rowCount = dataRange.rowCount;

      // Parse the input value
      const valueArray = value ? value.split(", ") : [];
      let updatesNeeded = false;

      // Process each row in chunks for better performance
      const CHUNK_SIZE = 50;
      for (let rowStart = 0; rowStart < rowCount; rowStart += CHUNK_SIZE) {
        const rowEnd = Math.min(rowStart + CHUNK_SIZE, rowCount);

        for (let i = rowStart; i < rowEnd; i++) {
          // Now access the local variable instead of the range property
          const rowValues = dataValues[i];
          if (!rowValues || rowValues.length <= Math.max(idColumnIndex, relatedFieldIndex)) {
            continue; // Skip invalid rows
          }

          const idColumnValue = String(rowValues[idColumnIndex] || "");
          const relatedFieldValue = String(rowValues[relatedFieldIndex] || "");

          if (!idColumnValue) {
            continue; // Skip rows without an ID
          }

          const relatedFieldValueArray = relatedFieldValue ? relatedFieldValue.split(", ") : [];

          // Case 1: Value is cleared and ID exists in related field - remove ID
          if (!value && relatedFieldValue && relatedFieldValueArray.includes(id)) {
            const newValueArray = relatedFieldValueArray.filter((v) => v !== id);
            const newValue = newValueArray.join(", ");
            dataRange.getCell(i, relatedFieldIndex).values = [[newValue]];
            updatesNeeded = true;
          }
          // Case 2: Value doesn't include current ID but related field includes the ID - remove ID
          else if (
            value &&
            !valueArray.includes(idColumnValue) &&
            relatedFieldValue &&
            relatedFieldValueArray.includes(id)
          ) {
            const newValueArray = relatedFieldValueArray.filter((v) => v !== id);
            const newValue = newValueArray.join(", ");
            dataRange.getCell(i, relatedFieldIndex).values = [[newValue]];
            updatesNeeded = true;
          }
          // Case 3: Value includes current ID but related field does not include the ID - add ID
          else if (
            value &&
            valueArray.includes(idColumnValue) &&
            (!relatedFieldValue || !relatedFieldValueArray.includes(id))
          ) {
            const newValueArray = relatedFieldValue ? relatedFieldValueArray : [];
            newValueArray.push(id);
            const newValue = newValueArray.join(", ");
            dataRange.getCell(i, relatedFieldIndex).values = [[newValue]];
            updatesNeeded = true;
          }
        }
      }

      // Only sync if changes were made
      if (updatesNeeded) {
        await context.sync();
      }
    } catch (error) {
      console.error(`Table ${relatedTableName} not found or other error, skipping: ${error}`);
    }
  } catch (error) {
    console.error(`Error in handleLinkedFieldsInRelatedSheet: ${error}`);
  }
}

/**
 * Validates that all values exist in a related sheet's @id column
 */
export async function checkIfAllValuesExistInRelatedSheet(
  context: Excel.RequestContext,
  relatedTableName: string,
  values: string
) {
  try {
    if (!values) {
      return true; // Empty values are always valid
    }

    const valuesArray = values.split(", ");
    if (valuesArray.length === 0) {
      return true;
    }

    try {
      const worksheet = context.workbook.worksheets.getItem(relatedTableName);
      const table = worksheet.tables.getItem(relatedTableName);
      const tableHeadersRange = table.getHeaderRowRange();
      tableHeadersRange.load("values");
      await context.sync();

      // Get the index of the @id field in the related table
      const idColumnIndex = tableHeadersRange.values[0].indexOf("@id");
      if (idColumnIndex === -1) {
        return false;
      }

      // For better performance with large tables, use getDataBodyRange
      const dataRange = table.getDataBodyRange();
      const idColumn = dataRange.getColumn(idColumnIndex);
      idColumn.load("values");
      await context.sync();

      // Create a Set for faster lookups
      const idSet = new Set();
      for (const row of idColumn.values) {
        if (row[0]) {
          idSet.add(String(row[0]));
        }
      }

      // Check all values exist in the Set
      for (const value of valuesArray) {
        if (value && !idSet.has(value)) {
          return false;
        }
      }

      return true;
    } catch (error) {
      console.error(`Error accessing table ${relatedTableName}: ${error}`);
      return false;
    }
  } catch (error) {
    console.error(`Error in checkIfAllValuesExistInRelatedSheet: ${error}`);
    return false;
  }
}

/**
 * Updates references when an ID is changed
 */
export async function updateRelatedFieldsValues(
  context: Excel.RequestContext,
  relatedTableName: string,
  relatedFieldName: string,
  oldValue: string,
  newValue: string
) {
  try {
    if (oldValue === newValue) {
      return;
    }

    // Get the table
    try {
      // Ensure the worksheet exists
      let worksheet;
      try {
        worksheet = context.workbook.worksheets.getItem(relatedTableName);
      } catch (err) {
        return;
      }

      // Ensure the table exists
      const table = worksheet.tables.getItem(relatedTableName);
      const tableHeadersRange = table.getHeaderRowRange();
      tableHeadersRange.load("values");
      await context.sync();

      // Get the index of the related field in the related table
      const relatedFieldIndex = tableHeadersRange.values[0].indexOf(relatedFieldName);
      if (relatedFieldIndex === -1) {
        return;
      }

      // Get data more efficiently by using the data body range
      const dataRange = table.getDataBodyRange();
      dataRange.load("values");
      dataRange.load("rowCount");
      await context.sync();

      // Store values in local variables
      const dataValues = dataRange.values;
      const rowCount = dataRange.rowCount;

      // Process rows and update references
      let updatesNeeded = false;

      for (let i = 0; i < rowCount; i++) {
        const rowValues = dataValues[i];
        if (!rowValues || rowValues.length <= relatedFieldIndex) {
          continue;
        }

        const cellValue = rowValues[relatedFieldIndex];
        const relatedFieldValue = cellValue !== undefined ? String(cellValue) : "";

        // Skip empty cells
        if (!relatedFieldValue) {
          continue;
        }

        // We need to check for exact match or as part of comma-separated list
        if (relatedFieldValue === oldValue) {
          // Direct replacement for exact match
          dataRange.getCell(i, relatedFieldIndex).values = [[newValue]];
          updatesNeeded = true;
        }
        // Check for value as part of comma-separated list
        else if (relatedFieldValue.includes(oldValue)) {
          const relatedFieldValueArray = relatedFieldValue.split(", ");

          // Only update if it's a complete match within the list
          if (relatedFieldValueArray.includes(oldValue)) {
            // Create a new array with the replaced value
            const newValueArray = relatedFieldValueArray.map((v) =>
              v === oldValue ? newValue : v
            );
            const updatedValue = newValueArray.join(", ");

            dataRange.getCell(i, relatedFieldIndex).values = [[updatedValue]];
            updatesNeeded = true;
          }
        }
      }

      // Only sync if we made changes
      if (updatesNeeded) {
        await context.sync();
      }
    } catch (error) {
      console.error(`Table ${relatedTableName} not found or other error: ${error}`);
    }
  } catch (error) {
    console.error(`Error in updateRelatedFieldsValues: ${error}`);
  }
}
