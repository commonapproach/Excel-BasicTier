import { IntlShape } from "react-intl";
import { TableInterface } from "../domain/interfaces/table.interface";
import {
  ModelType,
  SFFModelType,
  createInstance,
  map,
  mapSFFModel,
  predefinedCodeLists,
} from "../domain/models";
import { FieldType } from "../domain/models/Base";
import { validate } from "../domain/validation/validator";
import { createSFFModuleSheetsAndTables, createSheetsAndTables } from "../taskpane";
import { parseJsonLd } from "../utils/utils";

/* global Excel console */
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

    // Make sure we have data to import
    if (!filteredItems || filteredItems.length === 0) {
      throw new Error(
        intl.formatMessage({
          id: "import.messages.error.noValidData",
          defaultMessage: "No valid data to import. All records were filtered out.",
        })
      );
    }

    await importByData(intl, context, filteredItems);
  } catch (error: any) {
    console.error("Import failed:", error);
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
  await createSheetsAndTables(intl);

  // Check if data has any class from SFF module
  let needsSFFModuleTables = false;
  for (const data of jsonData) {
    if (data["@type"] && Object.keys(mapSFFModel).includes(data["@type"].split(":")[1])) {
      needsSFFModuleTables = true;
      break;
    }
  }

  if (needsSFFModuleTables) {
    await createSFFModuleSheetsAndTables(intl);
  }

  // Preload worksheets and tables to avoid repeated load operations
  const tableCache = new Map<
    string,
    {
      worksheet: Excel.Worksheet;
      table: Excel.Table;
      tableRange: Excel.Range;
      tableHeaders: string[];
      idColumnIndex: number;
    }
  >();

  // Group data by table for more efficient processing
  const dataByTable = groupDataByTable(jsonData);

  try {
    // Load all required tables in a single batch operation
    await loadTablesToCache(context, dataByTable, tableCache);

    // Process the data in correct order
    await processAllData(context, dataByTable, tableCache);
  } catch (error: any) {
    console.error("Error during data processing:", error);
    throw new Error(`Import failed: ${error.message || "Unknown error"}`);
  } finally {
    // Clean up waiting sheet
    try {
      const waitingSheet = context.workbook.worksheets.getItem(waitingSheetName);
      waitingSheet.delete();
      await context.sync();
    } catch (e) {
      console.warn("Could not remove waiting sheet:", e);
    }
  }
}

// New function to load tables with proper tracking
async function loadTablesToCache(
  context: Excel.RequestContext,
  dataByTable: Record<string, TableInterface[]>,
  tableCache: Map<string, any>
): Promise<void> {
  // Get unique table names from data plus code lists
  const uniqueTableNames = [...new Set([...Object.keys(dataByTable), ...predefinedCodeLists])];

  // First, check if all required tables exist
  const missingTables: string[] = [];
  for (const tableName of uniqueTableNames) {
    try {
      const worksheet = context.workbook.worksheets.getItem(tableName);
      worksheet.load("name");
      await context.sync();
    } catch (e) {
      if (!predefinedCodeLists.includes(tableName)) {
        missingTables.push(tableName);
      }
    }
  }

  if (missingTables.length > 0) {
    throw new Error(`Missing required tables: ${missingTables.join(", ")}`);
  }

  // Now load all tables with proper tracking
  for (const tableName of uniqueTableNames) {
    try {
      const worksheet = context.workbook.worksheets.getItem(tableName);
      // Track the worksheet explicitly
      context.trackedObjects.add(worksheet);

      worksheet.load("tables");
      await context.sync();

      try {
        const table = worksheet.tables.getItem(tableName);
        // Track the table explicitly
        context.trackedObjects.add(table);

        // Load table data with tracking
        const headerRange = table.getHeaderRowRange();
        context.trackedObjects.add(headerRange);
        headerRange.load("values");

        const tableRange = table.getRange();
        context.trackedObjects.add(tableRange);
        tableRange.load("values");
        await context.sync();

        const tableHeaders = headerRange.values[0];
        const idColumnIndex = tableHeaders.indexOf("@id");

        if (idColumnIndex === -1) {
          console.warn(`Table ${tableName} has no @id column!`);
          continue;
        }

        // Store in cache
        tableCache.set(tableName, {
          worksheet,
          table,
          tableRange,
          tableHeaders,
          idColumnIndex,
        });
      } catch (tableError) {
        console.warn(`Could not load table ${tableName}:`, tableError);
      }
    } catch (worksheetError) {
      console.warn(`Could not load worksheet ${tableName}:`, worksheetError);
    }
  }

  if (tableCache.size === 0) {
    throw new Error("Failed to load any tables. Import cannot continue.");
  }
}

// New function to process all data with better error handling
async function processAllData(
  context: Excel.RequestContext,
  dataByTable: Record<string, TableInterface[]>,
  tableCache: Map<string, any>
): Promise<void> {
  // Create record dependency map
  const dependencies = analyzeRecordDependencies(dataByTable);

  // Process tables in order of dependency complexity
  const processOrder = Object.keys(dataByTable).sort((a, b) => {
    const aDeps = dataByTable[a].reduce(
      (count, record) => count + (dependencies.get(record["@id"])?.size || 0),
      0
    );
    const bDeps = dataByTable[b].reduce(
      (count, record) => count + (dependencies.get(record["@id"])?.size || 0),
      0
    );
    return aDeps - bDeps;
  });

  // First pass: create all records with non-link fields
  for (const tableName of processOrder) {
    const tableData = dataByTable[tableName];
    const tableInfo = tableCache.get(tableName);

    if (!tableInfo) {
      console.warn(`No cache for table ${tableName}, skipping`);
      continue;
    }

    try {
      // Process non-link fields for this table
      await processSingleTableBasicFields(context, tableName, tableData, tableInfo);
    } catch (error) {
      console.error(`Error processing basic fields for ${tableName}:`, error);
      throw error;
    }
  }

  // Reload all tables after the first pass
  for (const [tableName, tableInfo] of tableCache.entries()) {
    try {
      // Create fresh range references
      const freshTable = tableInfo.worksheet.tables.getItem(tableName);
      context.trackedObjects.add(freshTable);

      const freshRange = freshTable.getRange();
      context.trackedObjects.add(freshRange);
      freshRange.load("values");

      await context.sync();

      // Update cache with fresh ranges
      tableInfo.table = freshTable;
      tableInfo.tableRange = freshRange;
    } catch (error) {
      console.warn(`Failed to reload table ${tableName}:`, error);
    }
  }

  // Second pass: process links between records
  try {
    await processAllLinks(context, dataByTable, tableCache);
  } catch (error) {
    console.error("Error processing links:", error);
    throw error;
  }

  // Final formatting
  for (const tableInfo of tableCache.values()) {
    try {
      tableInfo.table.getRange().format.autofitColumns();
      tableInfo.table.getRange().format.autofitRows();
    } catch (error) {
      console.warn("Failed to format table:", error);
    }
  }

  await context.sync();
}

// Process basic fields for a single table
async function processSingleTableBasicFields(
  context: Excel.RequestContext,
  tableName: string,
  tableData: TableInterface[],
  tableInfo: any
): Promise<void> {
  const { table, tableRange, tableHeaders, idColumnIndex } = tableInfo;
  const tableValues = tableRange.values;

  // Map existing IDs to row indices
  const idToRowIndex = new Map<string, number>();
  const existingIds = new Set<string>();
  for (let i = 1; i < tableValues.length; i++) {
    const rowId = tableValues[i][idColumnIndex]?.toString();
    if (rowId) {
      idToRowIndex.set(rowId, i);
      existingIds.add(rowId);
    }
  }

  // Find empty rows
  const emptyRows = [];
  for (let i = 1; i < tableValues.length; i++) {
    if (!tableValues[i][idColumnIndex]) {
      emptyRows.push(i);
    }
  }

  // Determine how many new rows we need
  const newRecords = tableData.filter((data) => !existingIds.has(data["@id"]));
  const rowsNeeded = newRecords.length - emptyRows.length;

  // Add rows if needed
  if (rowsNeeded > 0) {
    table.rows.add(null, rowsNeeded);

    // Reload table range after adding rows
    const freshRange = table.getRange();
    context.trackedObjects.add(freshRange);
    freshRange.load("values");
    await context.sync();

    const updatedValues = freshRange.values;
    // eslint-disable-next-line no-param-reassign
    tableInfo.tableRange = freshRange;

    // Update empty rows list with new rows
    for (let i = tableValues.length; i < updatedValues.length; i++) {
      emptyRows.push(i);
    }
  }

  // Process each record
  let emptyRowIndex = 0;
  let updatedCells = 0;

  for (const data of tableData) {
    const recordId = data["@id"];

    // Determine row for this record
    const rowIndex: number | undefined = idToRowIndex.get(recordId);
    let finalRowIndex: number; // This will hold our definite number value

    if (rowIndex === undefined) {
      if (emptyRowIndex >= emptyRows.length) {
        console.error(`No empty rows available for ${recordId} in ${tableName}`);
        continue;
      }

      finalRowIndex = emptyRows[emptyRowIndex++];

      // Set the ID in the new row
      const idCell = tableInfo.tableRange.getCell(finalRowIndex, idColumnIndex);
      idCell.values = [[recordId]];
      updatedCells++;

      // Update tracking
      idToRowIndex.set(recordId, finalRowIndex);
    } else {
      finalRowIndex = rowIndex;
    }

    // Process fields for this record
    const record = await processRecordFields(tableName, data);

    // Update cells with basic field values
    for (const [key, value] of Object.entries(record)) {
      const columnIndex = tableHeaders.indexOf(key);
      if (columnIndex !== -1) {
        try {
          const cell = tableInfo.tableRange.getCell(finalRowIndex, columnIndex);
          cell.values = [[value]];
          updatedCells++;
        } catch (error) {
          console.error(
            `Failed to update ${tableName} cell at ${finalRowIndex},${columnIndex}:`,
            error
          );
        }
      }
    }

    // Perform a sync every 100 cells to avoid overwhelming Excel
    if (updatedCells >= 100) {
      await context.sync();
      updatedCells = 0;
    }
  }

  // Final sync for this table
  if (updatedCells > 0) {
    await context.sync();
  }
}

// Process all links with improved duplicate detection
async function processAllLinks(
  context: Excel.RequestContext,
  dataByTable: Record<string, TableInterface[]>,
  tableCache: Map<string, any>
): Promise<void> {
  // Create a map of all imported records
  const importedRecordsMap = new Map<string, TableInterface>();
  Object.values(dataByTable)
    .flat()
    .forEach((record) => {
      if (record["@id"]) {
        importedRecordsMap.set(record["@id"], record);
      }
    });

  // Process links with smaller batches
  const BATCH_SIZE = 50; // Process 50 cells at a time
  let updateCount = 0;
  // Store table, row, column instead of cell references
  let pendingUpdates: Array<{
    tableRange: Excel.Range;
    rowIndex: number;
    columnIndex: number;
    value: string;
  }> = [];

  // Process links for each table
  for (const [tableName, tableData] of Object.entries(dataByTable)) {
    const tableInfo = tableCache.get(tableName);
    if (!tableInfo) continue;

    // Refresh the table range to ensure it's valid
    try {
      const freshTable = tableInfo.worksheet.tables.getItem(tableName);
      context.trackedObjects.add(freshTable);

      const freshRange = freshTable.getRange();
      context.trackedObjects.add(freshRange);
      freshRange.load("values");

      await context.sync();

      tableInfo.table = freshTable;
      tableInfo.tableRange = freshRange;
    } catch (error) {
      console.error(`Could not refresh table ${tableName}:`, error);
      continue;
    }

    const { tableRange, tableHeaders, idColumnIndex } = tableInfo;
    const tableValues = tableRange.values;

    // Map IDs to rows
    const idToRowIndex = new Map<string, number>();
    for (let i = 1; i < tableValues.length; i++) {
      const rowId = tableValues[i][idColumnIndex]?.toString();
      if (rowId) {
        idToRowIndex.set(rowId, i);
      }
    }

    // Process each record's links
    for (const data of tableData) {
      const recordId = data["@id"];
      const rowIndex = idToRowIndex.get(recordId);
      if (rowIndex === undefined) continue;

      const links = await extractAllLinks(tableName, data);

      // Process each link field
      for (const [fieldName, linkedIds] of Object.entries(links)) {
        if (!linkedIds || linkedIds.length === 0) continue;

        const columnIndex = tableHeaders.indexOf(fieldName);
        if (columnIndex === -1) continue;

        // Get model info to determine link type
        const cid = createInstance(tableName as ModelType | SFFModelType);
        let field;
        try {
          field = cid.getFieldByName(fieldName);
        } catch (e) {
          continue; // Skip if field not found
        }

        if (!field || field.type !== "link") continue;

        const isMultiLink = field.representedType === "array";
        const targetTable = field.link?.table?.className;

        // Validate links against both imported records and existing tables
        const validatedIds = targetTable
          ? validateLinks(linkedIds, targetTable, tableCache, importedRecordsMap)
          : linkedIds;

        if (validatedIds.length === 0) continue;

        // Get current cell value with improved parsing for existing links
        const currentValue = tableValues[rowIndex][columnIndex]?.toString() || "";
        // Normalize by splitting on commas and removing whitespace
        const currentIds = currentValue
          ? currentValue
              .split(",")
              .map((id: string) => id.trim())
              .filter(Boolean)
          : [];

        // Determine new value with enhanced duplicate detection
        let newValue;
        if (isMultiLink) {
          // Use Set operations to ensure all IDs are unique
          const combinedIds = [...new Set([...currentIds, ...validatedIds])];
          // Sort for consistent output (helps with future duplicate detection)
          combinedIds.sort();
          newValue = combinedIds.join(", ");
        } else {
          // For single links, use the existing or the first new valid link
          newValue = currentValue || validatedIds[0] || "";
        }

        // Only update if the value actually changed (to prevent unnecessary updates)
        if (newValue !== currentValue) {
          // Store reference information instead of the cell itself
          pendingUpdates.push({
            tableRange,
            rowIndex,
            columnIndex,
            value: newValue,
          });
          updateCount++;

          // Apply updates in batches
          if (updateCount >= BATCH_SIZE) {
            // Apply updates using stored references
            for (const update of pendingUpdates) {
              const cell = update.tableRange.getCell(update.rowIndex, update.columnIndex);
              context.trackedObjects.add(cell); // Track the cell explicitly
              cell.values = [[update.value]];
            }
            await context.sync();
            pendingUpdates = [];
            updateCount = 0;
          }
        }
      }
    }
  }

  // Apply any remaining updates
  if (pendingUpdates.length > 0) {
    try {
      for (const update of pendingUpdates) {
        const cell = update.tableRange.getCell(update.rowIndex, update.columnIndex);
        context.trackedObjects.add(cell); // Track the cell explicitly
        cell.values = [[update.value]];
      }
      await context.sync();
      pendingUpdates = [];
    } catch (error) {
      console.error("Failed to apply final batch of updates:", error);
    }
  }

  // Process bidirectional links in a separate pass
  await processBidirectionalLinks(context, dataByTable, tableCache, importedRecordsMap);
}

// Extract all links from a record (combining all extraction methods)
async function extractAllLinks(
  tableName: string,
  data: TableInterface
): Promise<Record<string, string[]>> {
  const directLinks = await extractDirectLinkFields(tableName, data);
  const nestedLinks = await extractNestedLinkFields(tableName, data);
  const recursiveLinks = await extractRecursiveLinkFields(tableName, data);

  // Combine all link sources
  const allLinks: Record<string, string[]> = {};

  [directLinks, nestedLinks, recursiveLinks].forEach((linkSource) => {
    Object.entries(linkSource).forEach(([fieldName, links]) => {
      if (!allLinks[fieldName]) {
        allLinks[fieldName] = [];
      }
      allLinks[fieldName] = [...new Set([...allLinks[fieldName], ...links])];
    });
  });

  return allLinks;
}

// Process bidirectional links with improved reliability and performance
async function processBidirectionalLinks(
  context: Excel.RequestContext,
  dataByTable: Record<string, TableInterface[]>,
  tableCache: Map<string, any>,
  importedRecordsMap: Map<string, TableInterface>
): Promise<void> {
  // Collect all bidirectional links
  const bidirectionalLinks: Map<
    string, // Target table
    Map<
      string, // Target record ID
      Map<
        string, // Target field
        { sourceIds: string[]; sourceTable: string } // Source record IDs and table
      >
    >
  > = new Map();

  // Collect all bidirectional links first
  for (const [tableName, tableData] of Object.entries(dataByTable)) {
    for (const data of tableData) {
      const recordId = data["@id"];
      const links = await extractAllLinks(tableName, data);

      // Get model info
      const cid = createInstance(tableName as ModelType | SFFModelType);

      // Process each link field to find bidirectional links
      for (const [fieldName, linkedIds] of Object.entries(links)) {
        if (!linkedIds || linkedIds.length === 0) continue;

        // Find field definition
        let field;
        try {
          field = cid.getFieldByName(fieldName);
        } catch (e) {
          continue; // Skip if field not found
        }

        if (!field || field.type !== "link" || !field.link?.table || !field.link.field) continue;

        const targetTable = field.link.table.className;
        const targetField = field.link.field;

        // Skip self-references
        if (targetTable === tableName) continue;

        // Add to bidirectional links map
        for (const targetId of linkedIds) {
          if (!bidirectionalLinks.has(targetTable)) {
            bidirectionalLinks.set(targetTable, new Map());
          }

          if (!bidirectionalLinks.get(targetTable)!.has(targetId)) {
            bidirectionalLinks.get(targetTable)!.set(targetId, new Map());
          }

          if (!bidirectionalLinks.get(targetTable)!.get(targetId)!.has(targetField)) {
            bidirectionalLinks.get(targetTable)!.get(targetId)!.set(targetField, {
              sourceIds: [],
              sourceTable: tableName,
            });
          }

          bidirectionalLinks
            .get(targetTable)!
            .get(targetId)!
            .get(targetField)!
            .sourceIds.push(recordId);
        }
      }
    }
  }

  // Process each target table in batches
  let updateCount = 0;
  // Store table, row, column instead of cell references
  let pendingUpdates: Array<{
    tableRange: Excel.Range;
    rowIndex: number;
    columnIndex: number;
    value: string;
  }> = [];
  const BATCH_SIZE = 50;

  for (const [targetTable, recordUpdates] of bidirectionalLinks.entries()) {
    const tableInfo = tableCache.get(targetTable);
    if (!tableInfo) continue;

    // Refresh the table range
    try {
      const freshTable = tableInfo.worksheet.tables.getItem(targetTable);
      context.trackedObjects.add(freshTable);

      const freshRange = freshTable.getRange();
      context.trackedObjects.add(freshRange);
      freshRange.load("values");

      await context.sync();

      tableInfo.table = freshTable;
      tableInfo.tableRange = freshRange;
    } catch (error) {
      console.error(`Could not refresh target table ${targetTable}:`, error);
      continue;
    }

    const { tableRange, tableHeaders, idColumnIndex } = tableInfo;
    const tableValues = tableRange.values;

    // Map IDs to rows
    const idToRowIndex = new Map<string, number>();
    for (let i = 1; i < tableValues.length; i++) {
      const rowId = tableValues[i][idColumnIndex]?.toString();
      if (rowId) {
        idToRowIndex.set(rowId, i);
      }
    }

    // Process each target record
    for (const [targetId, fieldUpdates] of recordUpdates.entries()) {
      const rowIndex = idToRowIndex.get(targetId);
      if (rowIndex === undefined) continue;

      // Process each field update
      for (const [targetField, { sourceIds, sourceTable }] of fieldUpdates.entries()) {
        const columnIndex = tableHeaders.indexOf(targetField);
        if (columnIndex === -1) continue;

        // Get target model info
        const targetCid = createInstance(targetTable as ModelType | SFFModelType);
        let field;
        try {
          field = targetCid.getFieldByName(targetField);
        } catch (e) {
          continue; // Skip if field not found
        }

        if (!field || field.type !== "link") continue;

        const isMultiLink = field.representedType === "array";

        // Validate source IDs
        const validatedSourceIds = validateLinks(
          sourceIds,
          sourceTable,
          tableCache,
          importedRecordsMap
        );
        if (validatedSourceIds.length === 0) continue;

        // Get current value with improved parsing
        const currentValue = tableValues[rowIndex][columnIndex]?.toString() || "";
        // Normalize by splitting on commas and removing whitespace
        const currentIds = currentValue
          ? currentValue
              .split(",")
              .map((id: string) => id.trim())
              .filter(Boolean)
          : [];

        // Determine new value with enhanced duplicate detection
        let newValue;
        if (isMultiLink) {
          // Use Set operations to ensure all IDs are unique
          const combinedIds = [...new Set([...currentIds, ...validatedSourceIds])];
          // Sort for consistent output (helps with future duplicate detection)
          combinedIds.sort();
          newValue = combinedIds.join(", ");
        } else {
          // For single links, use the existing or the first new valid link
          newValue = currentValue || validatedSourceIds[0] || "";
        }

        // Only update if value changed
        if (newValue !== currentValue) {
          // Store reference information instead of the cell itself
          pendingUpdates.push({
            tableRange,
            rowIndex,
            columnIndex,
            value: newValue,
          });
          updateCount++;

          if (updateCount >= BATCH_SIZE) {
            // Apply updates using stored references
            for (const update of pendingUpdates) {
              const cell = update.tableRange.getCell(update.rowIndex, update.columnIndex);
              context.trackedObjects.add(cell); // Track the cell explicitly
              cell.values = [[update.value]];
            }
            await context.sync();
            pendingUpdates = [];
            updateCount = 0;
          }
        }
      }
    }
  }

  // Apply any remaining updates
  if (pendingUpdates.length > 0) {
    try {
      for (const update of pendingUpdates) {
        const cell = update.tableRange.getCell(update.rowIndex, update.columnIndex);
        context.trackedObjects.add(cell); // Track the cell explicitly
        cell.values = [[update.value]];
      }
      await context.sync();
      pendingUpdates = [];
    } catch (error) {
      console.error("Failed to apply final batch of updates:", error);
    }
  }
}

// Extract direct link fields from the record
async function extractDirectLinkFields(
  tableName: string,
  data: TableInterface
): Promise<Record<string, string[]>> {
  const linkedFields: Record<string, string[]> = {};
  const cid = createInstance(tableName as ModelType | SFFModelType);

  // Get all link fields for this table
  const linkFields = cid.getAllFields().filter((field) => field.type === "link");

  for (const field of linkFields) {
    const fieldName = field.displayName || field.name;
    const fieldKey = field.name;
    let value;

    // Check main field name
    value = data[fieldKey];

    // Check for field name without namespace prefix
    if (!value && fieldKey.includes(":")) {
      const simpleName = fieldKey.split(":")[1];
      value = data[simpleName];
    }

    // If still no value found, try other versions of the field name
    if (!value) {
      // Try display name
      value = data[fieldName];

      // Try simple name if different from display name
      if (!value && fieldName !== fieldKey && fieldName !== fieldKey.split(":")[1]) {
        value = data[fieldName];
      }
    }

    if (value) {
      // Use the target table class for validation
      const processedLinks = handleLinkFields(value);
      if (processedLinks && processedLinks.length > 0) {
        linkedFields[fieldName] = processedLinks;
      }
    }
  }

  return linkedFields;
}

// Helper function to group data by table name for batch processing
function groupDataByTable(jsonData: TableInterface[]): Record<string, TableInterface[]> {
  const dataByTable: Record<string, TableInterface[]> = {};

  for (const data of jsonData) {
    if (!data["@type"]) continue;

    const tableName = data["@type"].split(":")[1];
    if (!dataByTable[tableName]) {
      dataByTable[tableName] = [];
    }

    dataByTable[tableName].push(data);
  }

  return dataByTable;
}

// Extract links from nested objects
async function extractNestedLinkFields(
  tableName: string,
  data: TableInterface
): Promise<Record<string, string[]>> {
  const linkedFields: Record<string, string[]> = {};
  const cid = createInstance(tableName as ModelType | SFFModelType);

  for (const field of cid.getAllFields()) {
    if (field.type === "object" && field.properties) {
      const fieldName = field.displayName || field.name;
      const fieldKey = field.name;

      let nestedData = data[fieldKey];

      // Check for field name without namespace prefix
      if (!nestedData && fieldKey.includes(":")) {
        const simpleName = fieldKey.split(":")[1];
        nestedData = data[simpleName];
      }

      // Skip if nestedData is not an object
      if (!nestedData || typeof nestedData !== "object" || Array.isArray(nestedData)) continue;

      // Process each property in the nested object
      for (const prop of field.properties) {
        if (prop.type === "link") {
          const propName = prop.displayName || prop.name;
          // Extract links from nested object
          let nestedValue;

          if (
            prop.name.includes(":") &&
            (nestedData as Record<string, any>)[prop.name.split(":")[1]] !== undefined
          ) {
            nestedValue = (nestedData as Record<string, any>)[prop.name.split(":")[1]];
          } else {
            nestedValue = (nestedData as Record<string, any>)[prop.name];
          }

          if (!nestedValue) continue;

          // Use the target table class for validation
          const processedLinks = handleLinkFields(nestedValue);
          if (processedLinks && processedLinks.length > 0) {
            // Use the parent field name + property name as the key
            linkedFields[`${fieldName}.${propName}`] = processedLinks;
          }
        }
      }
    }
  }

  return linkedFields;
}

// Recursively extract all linked fields with comprehensive checks
async function extractRecursiveLinkFields(
  tableName: string,
  data: TableInterface
): Promise<Record<string, string[]>> {
  const linkedFields: Record<string, string[]> = {};
  const cid = createInstance(tableName as ModelType | SFFModelType);

  // Recursively process all fields
  for (const [key, value] of Object.entries(data)) {
    if (key === "@type" || key === "@context" || !checkIfFieldIsRecognized(tableName, key)) {
      continue;
    }

    let fieldKey = key;

    // Try to find the full field key (including namespace if applicable)
    for (const field of cid.getAllFields()) {
      if (field.name.includes(":") && field.name.split(":")[1] === key) {
        fieldKey = field.name;
        break;
      }
    }

    const field = cid.getFieldByName(fieldKey);
    if (!field) continue;

    // Extract links recursively
    const links = processFieldForLinks(field, value);

    // Add any found links to our result
    for (const [linkName, linkIds] of Object.entries(links)) {
      if (!linkIds || linkIds.length === 0) continue;

      const fieldName = field.displayName || field.name;
      if (linkName === ".") {
        linkedFields[fieldName] = linkIds;
      } else {
        linkedFields[`${fieldName}${linkName}`] = linkIds;
      }
    }
  }

  return linkedFields;
}

// Helper function to process a field for links (moved outside the previous function)
function processFieldForLinks(field: FieldType, value: any): Record<string, string[]> {
  const result: Record<string, string[]> = {};

  if (field.type === "link") {
    // Direct link field
    const processedLinks = handleLinkFields(value);
    if (processedLinks && processedLinks.length > 0) {
      result["."] = processedLinks;
    }
  } else if (field.type === "object" && field.properties) {
    // Nested object with potential links
    if (typeof value === "object" && value !== null && !Array.isArray(value)) {
      for (const prop of field.properties) {
        const propValue =
          value[prop.name] || (prop.name.includes(":") ? value[prop.name.split(":")[1]] : null);

        if (propValue !== undefined && propValue !== null) {
          const nestedLinks = processFieldForLinks(prop, propValue);

          for (const [linkPath, linkIds] of Object.entries(nestedLinks)) {
            const propName = prop.displayName || prop.name;
            const fullPath = linkPath === "." ? `.${propName}` : `.${propName}${linkPath}`;
            result[fullPath] = linkIds;
          }
        }
      }
    }
  }

  return result;
}

// Existing helper functions
function handleLinkFields(value: any) {
  if (!value) return; // Skip if the value is empty

  // Convert to array if not already
  // eslint-disable-next-line no-param-reassign
  if (!Array.isArray(value)) value = [value];

  // Remove duplicates
  const uniqueValues = [...new Set(value as string[])];

  if (uniqueValues.length === 0) {
    return;
  }

  return uniqueValues;
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
        let warningMessage;
        if (Object.keys(map).includes(tableName)) {
          warningMessage = intl.formatMessage(
            {
              id: "import.messages.warning.unrecognizedField",
              defaultMessage:
                "Table <b>{tableName}</b> has unrecognized field <b>{fieldName}</b>. This field will be ignored.",
            },
            { tableName, fieldName: key, b: (str) => `<b>${str}</b>` }
          );
        } else {
          warningMessage = intl.formatMessage(
            {
              id: "import.messages.warning.notImport",
              defaultMessage:
                "Table <b>{tableName}</b> has unrecognized field <b>{fieldName}</b> and will not be imported.",
            },
            { tableName, fieldName: key, b: (str) => `<b>${str}</b>` }
          );
        }
        warnings.push(warningMessage);
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

// Enhanced version of processRecordFields for better field handling
async function processRecordFields(
  tableName: string,
  data: TableInterface
): Promise<Record<string, any>> {
  const record: Record<string, any> = {};
  const cid = createInstance(tableName as ModelType | SFFModelType);

  // First pass - process direct fields with proper name resolution
  for (const [key, value] of Object.entries(data)) {
    if (key === "@type" || key === "@context" || !checkIfFieldIsRecognized(tableName, key)) {
      continue;
    }

    // Resolve the full field key (including namespace if applicable)
    let fieldKey = key;
    for (const field of cid.getAllFields()) {
      if (field.name.includes(":") && field.name.split(":")[1] === key) {
        fieldKey = field.name;
        break;
      }
    }

    const field = cid.getFieldByName(fieldKey);
    if (!field) continue;

    // Skip link fields - they will be processed separately
    if (field.type === "link") continue;

    // Process by field type
    if (field.type === "object") {
      // Process nested object fields
      const objectFields = findLastFieldValueForNestedFields(data, field, {});
      Object.assign(record, objectFields);
    } else {
      // Process simple fields (string, number, boolean, select, multiselect)
      const fieldName = field.displayName || field.name;
      let newValue: any = value;

      // Handle select/multiselect fields
      if (newValue && (field.type === "select" || field.type === "multiselect")) {
        // Get options for this field
        let options: { id: string; name: string }[] = [];
        if (field.getOptionsAsync) {
          options = await field.getOptionsAsync();
        } else {
          options = field.selectOptions as { id: string; name: string }[];
        }

        // Process select fields - map ID to display name
        if (field.type === "select") {
          const optionValue = Array.isArray(newValue) ? newValue[0] : newValue;
          const optionField = options.find((opt) => opt.id === optionValue);
          newValue = optionField ? optionField.name : null;
        }
        // Process multiselect fields - map all IDs to display names
        else if (field.type === "multiselect" && Array.isArray(newValue)) {
          const mappedValues = newValue
            .map((val) => {
              const optionField = options.find((opt) => opt.id === val);
              return optionField ? optionField.name : null;
            })
            .filter(Boolean);

          newValue = mappedValues.length > 0 ? mappedValues.join(", ") : null;
        } else {
          newValue = null;
        }
      }

      // Handle boolean fields
      if (field.type === "boolean") {
        newValue = newValue === true ? "Yes" : "No";
      }

      // Convert non-primitive values to string if needed
      if (
        field.type !== "boolean" &&
        field.type !== "select" &&
        field.type !== "multiselect" &&
        newValue !== null &&
        newValue !== undefined
      ) {
        newValue = newValue.toString();
      }

      // Add field to record
      record[fieldName] = newValue;
    }
  }

  return record;
}

// Add this function to validate links against the actual table data
function validateLinks(
  linkedIds: string[],
  targetTable: string,
  tableCache: Map<string, any>,
  importedRecordsMap?: Map<string, TableInterface>
): string[] {
  // Create array for valid links
  const validLinks: string[] = [];
  const invalidLinks: string[] = [];

  // Check each linked ID
  for (const id of linkedIds) {
    let isValid = false;

    // First check if the ID exists in the current import batch
    if (importedRecordsMap && importedRecordsMap.has(id)) {
      const record = importedRecordsMap.get(id);
      // Verify the record is of the correct type
      if (record && record["@type"] && record["@type"].split(":")[1] === targetTable) {
        isValid = true;
      }
    }

    // If not found in import batch, check existing table data
    if (!isValid) {
      const tableInfo = tableCache.get(targetTable);
      if (tableInfo) {
        const { tableRange, idColumnIndex } = tableInfo;
        const tableValues = tableRange.values;

        // Search for ID in the table
        for (let i = 1; i < tableValues.length; i++) {
          const existingId = tableValues[i][idColumnIndex]?.toString();
          if (existingId === id) {
            isValid = true;
            break;
          }
        }
      } else if (predefinedCodeLists.includes(targetTable)) {
        // Special handling for code list tables
        isValid = true;
      }
    }

    // Add to appropriate list
    if (isValid) {
      validLinks.push(id);
    } else {
      invalidLinks.push(id);
    }
  }

  return validLinks;
}

// Add a new helper function to track dependencies between records in import batch
function analyzeRecordDependencies(
  dataByTable: Record<string, TableInterface[]>
): Map<string, Set<string>> {
  // Map from record ID to set of IDs it depends on
  const dependencies = new Map<string, Set<string>>();

  // Process each table for links
  for (const records of Object.values(dataByTable)) {
    for (const record of records) {
      if (!record["@id"]) continue;

      const recordId = record["@id"];
      if (!dependencies.has(recordId)) {
        dependencies.set(recordId, new Set());
      }

      // Check all fields for potential links
      for (const [field, value] of Object.entries(record)) {
        if (field === "@type" || field === "@context" || field === "@id") continue;

        // Check if value is a link or array of links
        if (typeof value === "string" && value.startsWith("esg://")) {
          dependencies.get(recordId)?.add(value);
        } else if (Array.isArray(value)) {
          for (const item of value) {
            if (typeof item === "string" && item.startsWith("esg://")) {
              dependencies.get(recordId)?.add(item);
            }
          }
        }
      }
    }
  }

  return dependencies;
}
