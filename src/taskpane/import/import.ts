import { IntlShape } from "react-intl";
import { getCodeListByTableName } from "../domain/fetchServer/getCodeLists";
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
import { getCidsTableSuffix } from "../utils/typeHelpers";
import { populateSeliGLI } from "../helpers/seliGLI";
import { populateSeliGLISFI } from "../helpers/seliGLISFI";
import {
  convertForFunderIdToForOrganization,
  convertIcAddressToPostalAddress,
  convertIcHasAddressToHasAddress,
  convertNumericalValueToHasNumericalValue,
  convertOrganizationIDFields,
  convertUnknownUnitToDescription,
  harmonizeCardinalityProperty,
  parseJsonLd,
} from "../utils/utils";

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

    // json-ld expansion/compaction
    // eslint-disable-next-line no-param-reassign
    jsonData = await parseJsonLd(jsonData);

    // Backward compatibility transformations (order matters)
    // 1. numerical_value -> hasNumericalValue
    // eslint-disable-next-line no-param-reassign
    jsonData = convertNumericalValueToHasNumericalValue(jsonData);

    // 2. Convert OrganizationID fields and normalize @type for backward compatibility
    // eslint-disable-next-line no-param-reassign
    jsonData = convertOrganizationIDFields(jsonData);

    // 3. unknown unit_of_measure -> unitDescription (async)
    const unitConversion = await convertUnknownUnitToDescription(jsonData);
    // eslint-disable-next-line no-param-reassign
    jsonData = unitConversion.data;
    const convertedUnknownUnits = unitConversion.converted;

    // 4. ic:hasAddress -> hasAddress
    const beforeAddressProp = JSON.stringify(jsonData);
    // eslint-disable-next-line no-param-reassign
    jsonData = convertIcHasAddressToHasAddress(jsonData);
    const propertyNamesConverted = JSON.stringify(jsonData) !== beforeAddressProp;

    // 5. Convert forFunderId to forOrganization for backward compatibility
    const originalDataFunding = JSON.stringify(jsonData);
    jsonData = convertForFunderIdToForOrganization(jsonData);
    const convertedFundingPropertyNames = JSON.stringify(jsonData) !== originalDataFunding;

    // 6. harmonize describesPopulation / i72:cardinality_of
    // eslint-disable-next-line no-param-reassign
    jsonData = harmonizeCardinalityProperty(jsonData);

    // 7. Convert legacy Address objects to PostalAddress/Address shape
    let convertedAddress = false;
    function convertAndTrack(obj: any): any {
      if (
        obj &&
        typeof obj === "object" &&
        obj["@type"] &&
        ((typeof obj["@type"] === "string" && obj["@type"].toLowerCase().includes("address")) ||
          (Array.isArray(obj["@type"]) &&
            obj["@type"].some((t: string) => t.toLowerCase().includes("address"))))
      ) {
        const original = JSON.stringify(obj);
        const converted = convertIcAddressToPostalAddress(obj);
        if (JSON.stringify(converted) !== original) convertedAddress = true;
        return converted;
      }
      return obj;
    }
    if (Array.isArray(jsonData)) {
      // eslint-disable-next-line no-param-reassign
      jsonData = jsonData.map(convertAndTrack);
    } else if (jsonData && typeof jsonData === "object") {
      // eslint-disable-next-line no-param-reassign
      jsonData = convertAndTrack(jsonData);
    }

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
    const { errors, warnings } = await validate(jsonData, "import", intl);

    // Add conversion warnings
    if (convertedAddress) {
      warnings.push(
        intl.formatMessage({
          id: "import.messages.warning.addressConverted",
          defaultMessage: "Some addresses were converted from the old format to the new format.",
        })
      );
    }
    if (propertyNamesConverted || convertedFundingPropertyNames) {
      warnings.push(
        intl.formatMessage({
          id: "import.messages.warning.propertyNamesConverted",
          defaultMessage:
            "Some property names were converted from old format (ic:hasAddress to hasAddress).",
        })
      );
    }
    if (convertedUnknownUnits) {
      warnings.push(
        intl.formatMessage({
          id: "import.messages.warning.unknownUnitsConverted",
          defaultMessage:
            "Some unknown unit_of_measure values were copied to unitDescription field. Please review and select the correct unit from the dropdown.",
        })
      );
    }

    // Check for modified code list items
    const codeListWarnings = await warnIfCodeListItemsModified(jsonData, intl);

    const allWarns = [
      ...warnings,
      ...warnIfUnrecognizedFieldsWillBeIgnored(jsonData, intl),
      ...codeListWarnings,
    ];

    allErrors = errors.join("<hr/>");
    allWarnings = allWarns.join("<hr/>");

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
      ? jsonData.filter((data: any) => {
          const suffix = getCidsTableSuffix(data["@type"]);
          return suffix ? Object.keys(fullMap).includes(suffix) : false;
        })
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
  try {
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
      const suffix = getCidsTableSuffix(data["@type"]);
      if (suffix && Object.keys(mapSFFModel).includes(suffix)) {
        needsSFFModuleTables = true;
        break;
      }
    }

    if (needsSFFModuleTables) {
      await createSFFModuleSheetsAndTables(intl);
    }
    const hasSeliGliRefs = jsonData.some((data: any) => {
      const forIndicator = data["forIndicator"];
      if (!forIndicator) return false;
      const values = Array.isArray(forIndicator) ? forIndicator : [forIndicator];
      return values.some((v: string) => typeof v === "string" && v.includes("codelist.commonapproach.org/SELI-GLI") );
});
    if (hasSeliGliRefs) {
    await populateSeliGLI();
}

    const hasSeliGliSfiRefs = jsonData.some((data: any) => {
      const forIndicator = data["forIndicator"];
      if (!forIndicator) return false;
      const values = Array.isArray(forIndicator) ? forIndicator : [forIndicator];
      return values.some(
        (v: string) => typeof v === "string" && v.includes("codelist.commonapproach.org/SELI-GLI-SFI")
      );
    });
    if (hasSeliGliSfiRefs) {
      await populateSeliGLISFI();
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
    }
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
  // Map<targetTable, Map<targetRecordId, Map<targetField, { sourceIds, sourceTable }>>>
  const recordUpdates: Map<
    string,
    Map<string, Map<string, { sourceIds: string[]; sourceTable: string }>>
  > = new Map();

  // Gather reverse link updates
  for (const [tableName, tableData] of Object.entries(dataByTable)) {
    const cid = createInstance(tableName as ModelType | SFFModelType);
    for (const data of tableData) {
      const recordId = data["@id"];
      if (!recordId) continue;

      // Use existing comprehensive link extraction
      const links = await extractAllLinks(tableName, data);

      for (const [fieldName, linkedIds] of Object.entries(links)) {
        if (!linkedIds || linkedIds.length === 0) continue;

        let fieldDef: any;
        try {
          fieldDef = cid.getFieldByName(fieldName);
        } catch (e) {
          continue;
        }
        if (!fieldDef || fieldDef.type !== "link" || !fieldDef.link?.table || !fieldDef.link.field)
          continue;

        const targetTable = fieldDef.link.table.className;
        const targetField = fieldDef.link.field;
        if (!targetTable || !targetField) continue;

        // Skip self references (Excel already handled forward link)
        if (targetTable === tableName) continue;

        for (const targetId of linkedIds) {
          if (!recordUpdates.has(targetTable)) recordUpdates.set(targetTable, new Map());
          const targetMap = recordUpdates.get(targetTable)!;
          if (!targetMap.has(targetId)) targetMap.set(targetId, new Map());
          const fieldMap = targetMap.get(targetId)!;
          if (!fieldMap.has(targetField))
            fieldMap.set(targetField, { sourceIds: [], sourceTable: tableName });
          const entry = fieldMap.get(targetField)!;
          if (!entry.sourceIds.includes(recordId)) entry.sourceIds.push(recordId);
        }
      }
    }
  }

  // Nothing to do
  if (recordUpdates.size === 0) return;

  // Batch update constants
  const BATCH_SIZE = 50;
  let updateCount = 0;
  let pendingUpdates: Array<{
    tableRange: Excel.Range;
    rowIndex: number;
    columnIndex: number;
    value: string;
  }> = [];

  // Apply reverse updates
  for (const [targetTable, targetRecords] of recordUpdates.entries()) {
    const tableInfo = tableCache.get(targetTable);
    if (!tableInfo) continue; // Target table not created/imported

    // Refresh table range
    try {
      const freshTable = tableInfo.worksheet.tables.getItem(targetTable);
      context.trackedObjects.add(freshTable);
      const freshRange = freshTable.getRange();
      context.trackedObjects.add(freshRange);
      freshRange.load("values");
      await context.sync();
      tableInfo.table = freshTable;
      tableInfo.tableRange = freshRange;
    } catch (e) {
      console.error(`Could not refresh target table ${targetTable}:`, e);
      continue;
    }

    const { tableRange, tableHeaders, idColumnIndex } = tableInfo;
    const tableValues = tableRange.values;
    const idToRowIndex = new Map<string, number>();
    for (let i = 1; i < tableValues.length; i++) {
      const rowId = tableValues[i][idColumnIndex]?.toString();
      if (rowId) idToRowIndex.set(rowId, i);
    }

    for (const [targetId, fieldsMap] of targetRecords.entries()) {
      const rowIndex = idToRowIndex.get(targetId);
      if (rowIndex === undefined) continue;

      for (const [targetField, { sourceIds, sourceTable }] of fieldsMap.entries()) {
        const columnIndex = tableHeaders.indexOf(targetField);
        if (columnIndex === -1) continue;

        // Confirm target field is link type
        let targetCidField;
        try {
          const targetCid = createInstance(targetTable as ModelType | SFFModelType);
          targetCidField = targetCid.getFieldByName(targetField);
        } catch (e) {
          continue;
        }
        if (!targetCidField || targetCidField.type !== "link") continue;
        const isMultiLink = targetCidField.representedType === "array";

        const validatedSourceIds = validateLinks(
          sourceIds,
          sourceTable,
          tableCache,
          importedRecordsMap
        );
        if (validatedSourceIds.length === 0) continue;

        const currentValue = tableValues[rowIndex][columnIndex]?.toString() || "";
        const currentIds = currentValue
          ? currentValue
              .split(",")
              .map((id: string) => id.trim())
              .filter(Boolean)
          : [];

        let newValue: string;
        if (isMultiLink) {
          const combined = [...new Set([...currentIds, ...validatedSourceIds])].sort();
          newValue = combined.join(", ");
        } else {
          newValue = currentValue || validatedSourceIds[0] || "";
        }

        if (newValue !== currentValue) {
          pendingUpdates.push({ tableRange, rowIndex, columnIndex, value: newValue });
          updateCount++;
          if (updateCount >= BATCH_SIZE) {
            for (const update of pendingUpdates) {
              const cell = update.tableRange.getCell(update.rowIndex, update.columnIndex);
              context.trackedObjects.add(cell);
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

  if (pendingUpdates.length > 0) {
    try {
      for (const update of pendingUpdates) {
        const cell = update.tableRange.getCell(update.rowIndex, update.columnIndex);
        context.trackedObjects.add(cell);
        cell.values = [[update.value]];
      }
      await context.sync();
    } catch (e) {
      console.error("Failed to apply final batch of reverse link updates:", e);
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

    const tableName = getCidsTableSuffix(data["@type"]);
    if (!tableName) continue;
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
    const suffixCheck = getCidsTableSuffix(data["@type"]);
    if (
      !data["@type"] ||
      !suffixCheck ||
      (!Object.keys(map).includes(suffixCheck) && !Object.keys(mapSFFModel).includes(suffixCheck))
    ) {
      continue;
    }
    const tableName = suffixCheck;
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

// Helper to normalize values for comparison
function normalizeValue(value: any): string {
  if (value === null || value === undefined) {
    return "";
  }
  if (Array.isArray(value)) {
    return value.map(normalizeValue).sort().join(",");
  }
  if (typeof value === "object") {
    return JSON.stringify(value);
  }
  return String(value).trim();
}

async function warnIfCodeListItemsModified(
  tableData: TableInterface[],
  intl: IntlShape
): Promise<string[]> {
  const warnings: string[] = [];

  for (const data of tableData) {
    const tableName = getCidsTableSuffix(data["@type"]) || "";

    // Only check predefined code list tables
    if (!predefinedCodeLists.includes(tableName)) {
      continue;
    }

    // Skip if no @id (validation will catch this elsewhere)
    if (!data["@id"]) {
      continue;
    }

    try {
      // Get the predefined code list for this table
      const codeList = await getCodeListByTableName(tableName);

      if (codeList && codeList.length > 0) {
        const existingItem = codeList.find((item) => item["@id"] === data["@id"]);

        if (existingItem) {
          // Check if imported data differs from predefined code list
          let hasChanges = false;

          // Cast to allow dynamic property access
          const existingItemObj = existingItem as unknown as Record<string, unknown>;
          for (const fieldName of Object.keys(existingItemObj)) {
            if (fieldName === "@id") continue; // Skip ID comparison

            const importedValue = normalizeValue(data[fieldName]);
            const codeListValue = normalizeValue(existingItemObj[fieldName]);

            if (importedValue !== codeListValue) {
              hasChanges = true;
              break;
            }
          }

          if (hasChanges) {
            warnings.push(
              intl.formatMessage(
                {
                  id: "import.messages.warning.codeListModified",
                  defaultMessage:
                    "Record with @id <b>{id}</b> in table <b>{tableName}</b> differs from the predefined code list item. The imported version will be used, which may cause inconsistencies.",
                },
                { id: data["@id"], tableName, b: (str) => `<b>${str}</b>` }
              )
            );
          }
        }
      }
    } catch (error) {
      // If we can't fetch the code list, just skip this check
      // eslint-disable-next-line no-console
      console.warn(`Could not check code list for ${tableName}:`, error);
    }
  }

  return warnings;
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
      const suffixCheck = getCidsTableSuffix(data["@type"]);
      if (
        !data["@type"] ||
        !suffixCheck ||
        (!Object.keys(map).includes(suffixCheck) && !Object.keys(mapSFFModel).includes(suffixCheck))
      ) {
        continue;
      }

      if (
        key === "@type" ||
        key === "@context" ||
        key === "@id" ||
        !checkIfFieldIsRecognized(getCidsTableSuffix(data["@type"])!, key)
      ) {
        continue;
      }

      const suffix = getCidsTableSuffix(data["@type"])!;
      const cid = Object.keys(map).includes(suffix)
        ? new map[suffix as ModelType]()
        : new mapSFFModel[suffix as SFFModelType]();

      const field = cid.getFieldByName(key);
      if (!field) continue;

      // OBJECT field normalization
      if (field.type === "object") {
        const fieldValue = handleNestedObjectFieldType(jsonData, field, value);
        if (fieldValue) data[key] = fieldValue;
        continue;
      }

      // LINK field normalization (object or array of objects/ids)
      if (field.type === "link") {
        if (value && typeof value === "object" && !Array.isArray(value)) {
          let id = (value as any)["@id"] as string | undefined;
          if (!id) {
            const nestedType = (value as any)["@type"];
            const nestedSuffix = nestedType
              ? getCidsTableSuffix(nestedType) || (typeof nestedType === "string" ? nestedType : "")
              : "";
            if (nestedSuffix) {
              id = (data["@id"] as string).replace(/\/$/, "") + "/" + nestedSuffix;
              (value as any)["@id"] = id;
            }
          }
          if (id) {
            data[key] = id;
            if ((value as any)["@type"]) {
              const alreadyExists = jsonData.some((d) => d && d["@id"] === id);
              if (!alreadyExists) jsonData.push(value as any);
            }
          } else {
            (data as any)[key] = undefined;
          }
          continue;
        }
        if (Array.isArray(value)) {
          const processedIds: string[] = [];
          for (const item of value) {
            if (typeof item === "object" && item !== null) {
              let id = (item as any)["@id"] as string | undefined;
              if (!id) {
                const nestedType = (item as any)["@type"];
                const nestedSuffix = nestedType
                  ? getCidsTableSuffix(nestedType) ||
                    (typeof nestedType === "string" ? nestedType : "")
                  : "";
                if (nestedSuffix) {
                  id = (data["@id"] as string).replace(/\/$/, "") + "/" + nestedSuffix;
                  (item as any)["@id"] = id;
                }
              }
              if (id) {
                processedIds.push(id);
                if ((item as any)["@type"]) {
                  const alreadyExists = jsonData.some((d) => d && d["@id"] === id);
                  if (!alreadyExists) jsonData.push(item as any);
                }
              }
            } else if (typeof item === "string") {
              processedIds.push(item);
            }
          }
          (data as any)[key] = [...new Set(processedIds)];
          continue;
        }
        continue; // primitive id string
      }

      // SELECT / MULTISELECT normalization when value supplied as object(s) {"@id": "..."}
      if (field.type === "select" || field.type === "multiselect") {
        if (value && typeof value === "object" && !Array.isArray(value)) {
          const id = (value as any)["@id"] as string | undefined;
          (data as any)[key] = id || undefined;
          continue;
        }
        if (Array.isArray(value)) {
          const ids: string[] = [];
          for (const item of value) {
            if (typeof item === "object" && item !== null) {
              const id = (item as any)["@id"] as string | undefined;
              if (id) ids.push(id);
            } else if (typeof item === "string") {
              ids.push(item);
            }
          }
          (data as any)[key] =
            field.type === "multiselect" ? [...new Set(ids)] : ids[0] || undefined;
          continue;
        }
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
    if (key === "@context" || !checkIfFieldIsRecognized(tableName, key)) {
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

      // Handle boolean fields (store localized Yes/No display)
      if (field.type === "boolean") {
        newValue = newValue === true ? "Yes" : "No";
      }

      // Handle number fields: coerce to numeric like Airtable (invalid => null)
      if (field.type === "number") {
        if (newValue !== null && newValue !== undefined && newValue !== "") {
          const parsed = Number(Array.isArray(newValue) ? newValue[0] : newValue);
          newValue = Number.isNaN(parsed) ? null : parsed;
        } else {
          newValue = null;
        }
      }

      // Convert remaining non-primitive values (excluding number/select/multiselect/boolean) to string
      if (
        field.type !== "boolean" &&
        field.type !== "select" &&
        field.type !== "multiselect" &&
        field.type !== "number" &&
        newValue !== null &&
        newValue !== undefined &&
        typeof newValue !== "string"
      ) {
        try {
          if (Array.isArray(newValue)) {
            newValue = newValue.join(", ");
          } else {
            newValue = newValue.toString();
          }
        } catch (_) {
          newValue = null;
        }
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
      if (record && record["@type"] && getCidsTableSuffix(record["@type"]) === targetTable) {
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
