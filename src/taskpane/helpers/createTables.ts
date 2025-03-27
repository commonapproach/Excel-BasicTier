/* global Excel, console */
import { IntlShape } from "react-intl";
import { createInstance, ModelType, SFFModelType } from "../domain/models";
import { Base } from "../domain/models/Base";

export async function createTable(
  context: Excel.RequestContext,
  sheetName: ModelType | SFFModelType,
  intl: IntlShape
) {
  try {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const tableName = sheetName.toString();

    // Pre-check: ensure no other worksheet holds a table with the standard table name.
    context.workbook.worksheets.load("items/name,tables/items/name");
    await context.sync();
    let conflictFound = false;
    for (const ws of context.workbook.worksheets.items) {
      if (ws.name !== sheetName) {
        ws.tables.items.forEach((tbl) => {
          if (tbl.name === tableName) {
            conflictFound = true;
          }
        });
        if (conflictFound) break;
      }
    }
    if (conflictFound) {
      throw new Error(
        intl.formatMessage(
          {
            id: "createTables.messages.error.tableNameMismatchWithSheet",
            defaultMessage:
              "A table with the standard name '{expectedTableName}' was found in a sheet with a different name. Please rename the table or the sheet to match.",
          },
          { expectedTableName: tableName, b: (str) => `<b>${str}</b>` }
        )
      );
    }

    // Ensure @id is the first field in the array by sorting
    let fields = getTableFieldsNames(sheetName);
    const idIndex = fields.findIndex((field) => field === "@id");
    if (idIndex > 0) {
      // If @id exists but isn't first, move it to the beginning
      fields = ["@id", ...fields.filter((f) => f !== "@id")];
    } else if (idIndex === -1) {
      // If @id doesn't exist, add it
      fields = ["@id", ...fields];
    }

    const headers = fields.map((field) => (field === "@id" ? "'@id" : field)); // eslint-disable-line quotes
    const DEFAULT_ROW_COUNT = 1000; // Define default row count constant

    // Check if the sheet already contains ANY table (not just one with our expected name)
    sheet.load("tables/items");
    sheet.load("tables/items/length");
    sheet.load("tables/items/name");
    await context.sync();

    let existingTables = sheet.tables.items;
    let table: Excel.Table | null = null;
    let tableExists = false;

    // Conflict check: on the standard sheet a table with the expected name must exist.
    if (existingTables.length > 0) {
      let standardTable: Excel.Table | null = null;
      for (const t of existingTables) {
        if (t.name === tableName) {
          standardTable = t;
          break;
        }
      }
      if (!standardTable) {
        throw new Error(
          intl.formatMessage(
            {
              id: "createTables.messages.error.tableNameMismatch",
              defaultMessage:
                "Mismatch between sheet name and table name. Expected table name: '{expectedTableName}'.",
            },
            { expectedTableName: tableName, b: (str) => `<b>${str}</b>` }
          )
        );
      } else {
        table = standardTable;
        tableExists = true;
      }
    }

    // If any tables exist on the sheet, we'll work with the first one
    if (existingTables.length > 0) {
      try {
        // First, try to get our specifically named table
        table = sheet.tables.getItem(tableName);
        tableExists = true;
      } catch (error) {
        // Our specific table doesn't exist, but there is another table
        // Let's use that one and rename it
        table = existingTables[0];
        tableExists = true;

        // Add to tracked objects and rename
        context.trackedObjects.add(table);
        table.name = tableName;
        await context.sync();
      }

      // At this point we have a table to work with - get its headers
      const tableHeaders = table.getHeaderRowRange();
      context.trackedObjects.add(tableHeaders);
      tableHeaders.load("values");
      table.load("columns/count");
      await context.sync();

      let existingHeaders = tableHeaders.values[0];

      // Check for any duplicate @id columns that may exist from previous operations
      try {
        const idColumns = [];
        for (let i = 0; i < existingHeaders.length; i++) {
          const header = existingHeaders[i];
          if (
            header === "@id" ||
            header === "'@id" ||
            (typeof header === "string" && header.startsWith("@id") && /^@id\d*$/.test(header))
          ) {
            idColumns.push(i);
          }
        }

        // If we have multiple @id columns, delete all except the first one we found
        if (idColumns.length > 1) {
          // Sort in reverse order so we delete from right to left (avoiding index shifts)
          idColumns.sort((a, b) => b - a);

          // Keep the first one we found, delete the rest
          for (let i = 0; i < idColumns.length - 1; i++) {
            table.columns.getItem(idColumns[i]).delete();
          }

          // Refresh our headers after deletion
          const refreshedHeaders = table.getHeaderRowRange();
          context.trackedObjects.add(refreshedHeaders);
          refreshedHeaders.load("values");
          await context.sync();

          existingHeaders = refreshedHeaders.values[0];
        }
      } catch (duplicateError) {
        console.warn(`Error checking for duplicate @id columns: ${duplicateError}`);
      }

      // Check if @id column exists anywhere in the table
      const idColumnExists = existingHeaders.some((h) => h === "@id" || h === "'@id");

      // If @id doesn't exist anywhere, create it as the first column
      if (!idColumnExists) {
        try {
          console.log("Adding missing @id column at position 0");

          // Insert a new column at position 0
          const newColumn = table.columns.add(0);
          context.trackedObjects.add(newColumn);
          await context.sync();

          // Refresh our headers after adding the new column
          let updatedHeaders = table.getHeaderRowRange();
          context.trackedObjects.add(updatedHeaders);
          updatedHeaders.load("values");
          await context.sync();

          // Set the header value for the new column
          const newHeaderCell = newColumn.getHeaderRowRange();
          newHeaderCell.values = [["'@id"]];
          await context.sync();

          // Refresh our headers after changing the new column header
          updatedHeaders = table.getHeaderRowRange();
          context.trackedObjects.add(updatedHeaders);
          updatedHeaders.load("values");
          await context.sync();

          // Update existingHeaders for later processing
          existingHeaders = updatedHeaders.values[0];
        } catch (addIdError) {
          console.error(`Error adding @id column: ${addIdError}`);
        }
      }

      // MODIFICATION: Create a map of our standard headers for easier reference
      const standardHeadersMap = new Map<string, number>();
      headers.forEach((header, index) => {
        standardHeadersMap.set(header === "'@id" ? "@id" : header, index);
      });

      // MODIFICATION: Instead of completely restructuring, only update headers we care about
      // This preserves any custom columns the user may have added
      for (let i = 0; i < existingHeaders.length; i++) {
        const existingHeader = existingHeaders[i];
        // Only modify headers that are part of our standard set
        if (standardHeadersMap.has(existingHeader)) {
          const standardIndex = standardHeadersMap.get(existingHeader);
          const standardHeader = headers[standardIndex!];
          // Update if the format is different (e.g., '@id' vs "'@id")
          if (existingHeader !== standardHeader) {
            tableHeaders.getCell(0, i).values = [[standardHeader]];
          }
          // Mark this header as processed
          standardHeadersMap.delete(existingHeader);
        }
      }

      // NEW: Add missing standard columns to the table if they don't exist
      if (standardHeadersMap.size > 0) {
        try {
          // Get current table range for reference
          const tableRange = table.getRange();
          tableRange.load("columnCount, address");
          await context.sync();

          // Add each missing column
          const missingHeaders = Array.from(standardHeadersMap.entries());
          for (const [header, originalIndex] of missingHeaders) {
            try {
              // Add a new column to the table
              const newColumn = table.columns.add();
              context.trackedObjects.add(newColumn);
              await context.sync();

              // Set the header value
              const headerRange = newColumn.getHeaderRowRange();
              context.trackedObjects.add(headerRange);
              headerRange.values = [[header === "@id" ? "'@id" : header]];

              // Apply the same header styling
              headerRange.format.fill.color = "#3CAEA3";
              headerRange.format.font.color = "#000000";
              headerRange.format.font.bold = true;
              await context.sync();

              // Clean up tracked object
              context.trackedObjects.remove(newColumn);
            } catch (columnError) {
              console.error(`Error adding column ${header}: ${columnError}`);
            }
          }
        } catch (addColumnError) {
          console.error(`Error adding missing columns: ${addColumnError}`);
        }
      }

      // Improved row checking - handle empty tables better
      try {
        // First check if the table has any rows
        table.rows.load("count");
        await context.sync();

        // Fix row addition issue by checking different ways to add rows
        if (table.rows.count > 0) {
          // Table has rows, check the data range
          try {
            const dataRange = table.getDataBodyRange();
            dataRange.load("rowCount");
            await context.sync();

            // Add rows if needed to match the default row count
            if (dataRange.rowCount < DEFAULT_ROW_COUNT) {
              const rowsToAdd = DEFAULT_ROW_COUNT - dataRange.rowCount;
              if (rowsToAdd > 0) {
                // Try all these approaches in sequence until one works
                try {
                  for (let i = 0; i < rowsToAdd; i++) {
                    table.rows.add(-1); // Use -1 to consistently add at the end
                    // Sync every 100 rows to avoid overloading
                    if (i % 100 === 0) {
                      await context.sync();
                    }
                  }
                  await context.sync();
                } catch (error) {
                  console.error(`All methods to add rows failed: ${error}`);
                }
              }
            }
          } catch (rangeError) {
            console.warn(`Error checking data range: ${rangeError}`);
            // Try adding rows directly
            try {
              // Add in batches of 100 to avoid timeout
              const batchSize = 100;
              const batches = Math.ceil(DEFAULT_ROW_COUNT / batchSize);

              for (let i = 0; i < batches; i++) {
                const rowsToAdd = Math.min(batchSize, DEFAULT_ROW_COUNT - i * batchSize);
                if (rowsToAdd > 0) {
                  table.rows.add(-1, rowsToAdd); // Use -1 instead of undefined
                  await context.sync();
                }
              }
            } catch (error) {
              console.error(`Failed to add rows in batches: ${error}`);
            }
          }
        } else {
          // Table is empty, add the default number of rows in batches
          const batchSize = 100;
          const batches = Math.ceil(DEFAULT_ROW_COUNT / batchSize);

          for (let i = 0; i < batches; i++) {
            try {
              const rowsToAdd = Math.min(batchSize, DEFAULT_ROW_COUNT - i * batchSize);
              if (rowsToAdd > 0) {
                table.rows.add(-1, rowsToAdd); // Use -1 instead of undefined
                await context.sync();
              }
            } catch (error) {
              console.warn(`Failed to add batch ${i}: ${error}`);
            }
          }
        }
      } catch (rowError) {
        console.warn(`Error checking table rows: ${rowError}`);
      }

      await context.sync();
    } else {
      // No existing table - create from scratch with @id as the first column
      tableExists = false;
    }

    if (!tableExists) {
      try {
        // Create table and apply all formatting in one batch
        const rangeAddress = `A1:${String.fromCharCode(65 + headers.length - 1)}1`;
        const tableRange = sheet.getRange(rangeAddress);
        tableRange.values = [headers];

        // Create the table and track it immediately
        table = sheet.tables.add(tableRange.getResizedRange(DEFAULT_ROW_COUNT, 0), true);
        context.trackedObjects.add(table);
        table.name = tableName;
        table.showTotals = false;

        // Apply immediate formatting then sync
        sheet.freezePanes.freezeRows(1);
        await context.sync();

        // Get a fresh reference after table creation
        table = sheet.tables.getItem(tableName);
        context.trackedObjects.add(table);

        // Apply remaining formatting
        const tableRangeFormatting = table.getRange();
        context.trackedObjects.add(tableRangeFormatting);
        tableRangeFormatting.format.wrapText = true;
        tableRangeFormatting.format.verticalAlignment = "Center";

        const headerRangeForWidth = table.getHeaderRowRange();
        context.trackedObjects.add(headerRangeForWidth);
        headerRangeForWidth.format.columnWidth = 220;

        tableRangeFormatting.format.autofitColumns();
        tableRangeFormatting.format.autofitRows();

        // Sync after basic table creation and format
        await context.sync();
      } catch (createError) {
        console.error(`Error creating table ${tableName}: ${createError}`);
        throw new Error(
          intl.formatMessage(
            {
              id: "createTables.messages.error.tableCreationFailed",
              defaultMessage: "Failed to create table '{tableName}': {error}.",
            },
            {
              tableName,
              error: String(createError),
              b: (str) => `<b>${str}</b>`,
            }
          )
        );
      }
    }

    // Get a fresh reference to the table before styling to ensure it's valid
    try {
      if (table) context.trackedObjects.remove(table);
      table = sheet.tables.getItem(tableName);
      context.trackedObjects.add(table);
      await context.sync();
    } catch (error) {
      console.error(`Failed to get fresh table reference for ${tableName}: ${error}`);
      throw new Error(
        intl.formatMessage(
          {
            id: "createTables.messages.error.tableRefreshFailed",
            defaultMessage: "Failed to refresh table '{tableName}': {error}.",
          },
          {
            tableName,
            error: String(error),
            b: (str) => `<b>${str}</b>`,
          }
        )
      );
    }

    // Apply styling to table with extra precautions
    try {
      // Set table style first
      table.style = "TableStyleLight1";
      await context.sync();

      // Get fresh header range with tracking for styling
      const headerRange = table.getHeaderRowRange();
      context.trackedObjects.add(headerRange);
      headerRange.format.fill.color = "#3CAEA3";
      headerRange.format.font.color = "#000000";
      headerRange.format.font.bold = true;
      await context.sync();

      // Only now load header values for linked field identification
      const headerRangeForValues = table.getHeaderRowRange();
      context.trackedObjects.add(headerRangeForValues);
      headerRangeForValues.load("values");
      await context.sync();

      // Style linked fields
      const model = createInstance(sheetName);
      const linkedFields = model.getAllFields().filter((field) => field.link);

      if (linkedFields.length > 0) {
        // Get data range for styling (if any data exists)
        try {
          const dataBodyRange = table.getDataBodyRange();
          context.trackedObjects.add(dataBodyRange);
          dataBodyRange.load("rowCount");
          await context.sync();

          if (dataBodyRange.rowCount > 0) {
            const allHeaders = headerRangeForValues.values[0];
            for (const field of linkedFields) {
              const fieldName = field.displayName || field.name;
              const columnIndex = allHeaders.indexOf(fieldName);

              if (columnIndex !== -1) {
                try {
                  const columnRange = dataBodyRange.getColumn(columnIndex);
                  context.trackedObjects.add(columnRange);
                  columnRange.format.font.color = "#666666";
                  await context.sync();
                } catch (columnError) {
                  console.error(`Error styling column for ${fieldName}: ${columnError}`);
                  throw new Error(
                    intl.formatMessage(
                      {
                        id: "createTables.messages.warning.tableStylesFailed",
                        defaultMessage: "Table was created but styling failed: {error}.",
                      },
                      {
                        error: String(columnError),
                        b: (str) => `<b>${str}</b>`,
                      }
                    )
                  );
                }
              }
            }
          }
        } catch (dataBodyError) {
          console.warn(`Cannot style data rows in ${tableName}: Empty table or ${dataBodyError}`);
          throw new Error(
            intl.formatMessage(
              {
                id: "createTables.messages.warning.tableStylesFailed",
                defaultMessage: `Table was created but styling failed: ${dataBodyError}`,
              },
              { error: String(dataBodyError) }
            )
          );
        }
      }

      // Apply checkbox formatting to boolean fields
      await applyCheckboxesToBooleanFields(
        context,
        table,
        model,
        headerRangeForValues.values[0],
        intl
      );

      // After the linked fields styling block in createTable:
      try {
        const dataBodyRange = table.getDataBodyRange();
        dataBodyRange.load("values/length");
        await context.sync();

        const headerRangeForValues = table.getHeaderRowRange();
        headerRangeForValues.load("values");
        await context.sync();

        const headers = headerRangeForValues.values[0];
        const idColIndex = headers.findIndex((h) => h === "@id" || h === "'@id");
        if (idColIndex !== -1 && dataBodyRange.values.length > 0) {
          const idColumnRange = dataBodyRange.getColumn(idColIndex);
          idColumnRange.format.font.color = "#666666";
          await context.sync();
        }
      } catch (idStyleError) {
        console.error("Error styling @id column:", idStyleError);
        throw new Error(
          intl.formatMessage(
            {
              id: "createTables.messages.warning.tableStylesFailed",
              defaultMessage: "Table was created but styling failed: {error}.",
            },
            {
              error: String(idStyleError),
              b: (str) => `<b>${str}</b>`,
            }
          )
        );
      }
    } catch (error) {
      console.error(`Error applying styles to table ${tableName}: ${error}`);
      throw new Error(
        intl.formatMessage(
          {
            id: "createTables.messages.warning.tableStylesFailed",
            defaultMessage: `Table was created but styling failed: ${error}`,
          },
          { error: String(error) }
        )
      );
    } finally {
      // Clean up tracked objects
      if (table) {
        try {
          context.trackedObjects.remove(table);
        } catch (removeError) {
          console.error(`Error removing table from tracked objects: ${removeError}`);
        }
      }
    }
  } catch (outerError: any) {
    // MODIFICATION: Suppress error messages for table creation when they're likely due to custom columns
    if (
      outerError.toString().includes("dimensions of the range") ||
      outerError.toString().includes("doesn't match the size")
    ) {
      console.log(`Notice: Table ${sheetName} has custom columns - continuing without error`);
    } else {
      console.error(`Fatal error processing table ${sheetName}: ${outerError}`);
      // If it's already an intl-formatted error, just rethrow it
      if (
        outerError instanceof Error &&
        !outerError.message.includes("Critical error creating table")
      ) {
        throw outerError;
      } else {
        throw new Error(
          intl.formatMessage(
            {
              id: "createTables.messages.error.tableFatalError",
              defaultMessage: "Critical error creating table '{sheetName}': {error}.",
            },
            {
              sheetName: String(sheetName),
              error: String(outerError),
              b: (str) => `<b>${str}</b>`,
            }
          )
        );
      }
    }
  }
}

async function applyCheckboxesToBooleanFields(
  context: Excel.RequestContext,
  table: Excel.Table,
  model: Base,
  headers: string[],
  intl: IntlShape
) {
  try {
    // Get boolean fields from the model
    const booleanFields = model.getAllFields().filter((field) => field.type === "boolean");

    if (booleanFields.length === 0) {
      return; // No boolean fields to process
    }

    // Get data body range
    try {
      const dataBodyRange = table.getDataBodyRange();
      context.trackedObjects.add(dataBodyRange);
      dataBodyRange.load("rowCount");
      await context.sync();

      if (dataBodyRange.rowCount === 0) {
        return; // No data rows to format
      }

      // Process each boolean field
      for (const field of booleanFields) {
        const fieldName = field.displayName || field.name;
        const columnIndex = headers.indexOf(fieldName);

        if (columnIndex === -1) {
          continue; // Field not found in headers
        }

        try {
          const columnRange = dataBodyRange.getColumn(columnIndex);
          context.trackedObjects.add(columnRange);

          // Set Yes/No dropdown validation
          columnRange.dataValidation.clear();
          columnRange.dataValidation.rule = {
            list: {
              inCellDropDown: true,
              source: "Yes,No",
            },
          };

          // Set default value to "No" for empty cells and convert existing values
          columnRange.load("values");
          await context.sync();

          // Convert existing values to Yes/No format
          const values = columnRange.values;
          for (let i = 0; i < values.length; i++) {
            const currentValue = values[i][0];

            if (currentValue === "") {
              // Set empty cells to "No"
              dataBodyRange.getCell(i, columnIndex).values = [["No"]];
            } else if (typeof currentValue === "boolean") {
              // Convert boolean to Yes/No
              dataBodyRange.getCell(i, columnIndex).values = [[currentValue ? "Yes" : "No"]];
            } else if (typeof currentValue === "string") {
              // Convert string "true"/"false" (or variations) to Yes/No
              const lowerValue = currentValue.toLowerCase();
              if (lowerValue === "true" || lowerValue === "yes") {
                dataBodyRange.getCell(i, columnIndex).values = [["Yes"]];
              } else if (lowerValue === "false" || lowerValue === "no") {
                dataBodyRange.getCell(i, columnIndex).values = [["No"]];
              } else {
                // For any other string, default to "No"
                dataBodyRange.getCell(i, columnIndex).values = [["No"]];
              }
            } else {
              // For any non-boolean, non-string value, default to "No"
              dataBodyRange.getCell(i, columnIndex).values = [["No"]];
            }
          }

          await context.sync();
        } catch (columnError) {
          console.error(`Error applying Yes/No dropdown to column ${fieldName}: ${columnError}`);
          throw new Error(
            intl.formatMessage(
              {
                id: "createTables.messages.error.populateCodeList",
                defaultMessage: "Error populating code list for table <b>{tableName}</b>.",
              },
              {
                tableName: table.name,
                fieldName,
                b: (str) => `<b>${str}</b>`,
              }
            )
          );
        }
      }
    } catch (rangeError) {
      console.warn(`Cannot apply Yes/No dropdowns: ${rangeError}`);
      throw new Error(
        intl.formatMessage(
          {
            id: "createTables.messages.error.populateCodeList",
            defaultMessage: `Error applying Yes/No dropdowns: ${rangeError}`,
          },
          { error: String(rangeError) }
        )
      );
    }
  } catch (error) {
    console.error(`Error in applyCheckboxesToBooleanFields: ${error}`);
    throw new Error(
      intl.formatMessage(
        {
          id: "createTables.messages.error.populateCodeList",
          defaultMessage: "Error populating code list for table <b>{tableName}</b>.",
        },
        {
          tableName: table.name,
          b: (str) => `<b>${str}</b>`,
        }
      )
    );
  }
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
