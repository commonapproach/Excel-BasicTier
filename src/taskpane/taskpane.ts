/* global Office Excel setTimeout clearTimeout console */
import { IntlShape } from "react-intl";
import { dialogHandler } from "./context/DialogContext";
import { getCodeListByTableName } from "./domain/fetchServer/getCodeLists";
import { getUnitOptions } from "./domain/fetchServer/getUnitsOfMeasure";
import {
  createInstance,
  ignoredFields,
  map,
  mapSFFModel,
  ModelType,
  predefinedCodeLists,
  SFFModelType,
} from "./domain/models";
import { createHiddenTables } from "./helpers/createHiddenTables";
import { createTable } from "./helpers/createTables";
import {
  isCellInRange as importedCellInRange,
  processMultiSelectValue,
  showDialog,
} from "./helpers/excelUtilities";
import {
  checkIfAllValuesExistInRelatedSheet,
  handleLinkedFieldsInRelatedSheet,
  updateRelatedFieldsValues,
} from "./helpers/handleLinkedFieldsOnOtherSheet";

// Hidden sheets hold code lists used for data validation. StreetType & StreetDirection removed; UnitsOfMeasure added.
const hiddenSheets = ["ProvinceTerritory", "OrganizationType", "Locality", "UnitsOfMeasure"];

// Replace existing silent mode implementation with this more comprehensive solution
let inSilentMode = false;
let silentModeOperations = 0;
let pendingOperations = false; // New flag to track operations in progress
const registeredHandlers = new Map<string, { remove: () => void }>();
const SILENT_MODE_COOLDOWN = 2000; // 2 seconds cooldown after operations
const HANDLER_REGISTRATION_DELAY = SILENT_MODE_COOLDOWN * 1.5; // 3 seconds for event handler registration

// Add debounce utility at the top with other utility functions
const debounceTimers = new Map<string, ReturnType<typeof setTimeout>>();

function debounce(key: string, func: Function, wait: number) {
  // Cancel previous timer for this key if it exists
  if (debounceTimers.has(key)) {
    clearTimeout(debounceTimers.get(key)!);
    debounceTimers.delete(key);
  }

  // Set new timer
  const timer = setTimeout(() => {
    func();
    debounceTimers.delete(key);
  }, wait);

  debounceTimers.set(key, timer);
}

// Enhanced silent mode with more aggressive protection
function enableSilentMode() {
  silentModeOperations++;
  inSilentMode = true;
  pendingOperations = true;
}

function disableSilentMode() {
  silentModeOperations = Math.max(0, silentModeOperations - 1);
  inSilentMode = silentModeOperations > 0;

  if (silentModeOperations === 0) {
    // Add cooldown period to prevent immediate event firing
    setTimeout(() => {
      pendingOperations = false;
    }, SILENT_MODE_COOLDOWN);
  }
}

// Remove all existing event handlers before adding new ones
async function cleanupAllEventHandlers(context: Excel.RequestContext) {
  try {
    // First, remove any registered handlers we're tracking
    for (const [key, handler] of registeredHandlers.entries()) {
      try {
        handler.remove();
      } catch (e) {
        console.error(`Failed to remove handler ${key}: ${e}`);
      }
    }
    registeredHandlers.clear();

    // Get all worksheets
    const sheets = context.workbook.worksheets;
    sheets.load("items");
    await context.sync();

    await context.sync();
  } catch (error) {
    console.error("Error in cleanupAllEventHandlers:", error);
  }
}

Office.onReady(() => {
  // If needed, Office.js is ready to be called.

  // Add lookup multi-select functionality if the user have all the standard tables
  if (Office.context.host === Office.HostType.Excel) {
    addLookupMultiSelectHandlerToAllTables();
  }
});

export async function createSheetsAndTables(intl: IntlShape) {
  enableSilentMode(); // Enable silent mode before operations
  try {
    await Excel.run(async (context) => {
      try {
        // Load all worksheets in a single batch operation
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        // Get existing sheet names to avoid recreation attempts
        const existingSheetNames = new Set(sheets.items.map((sheet) => sheet.name));

        // Create all needed sheets in batch
        const sheetsToAdd = Object.keys(map).filter((name) => !existingSheetNames.has(name));

        // Create any missing hidden sheets (now includes UnitsOfMeasure)
        const hiddenSheetsToAdd = hiddenSheets.filter((name) => !existingSheetNames.has(name));

        // Add all sheets in one batch
        for (const sheetName of sheetsToAdd) {
          sheets.add(sheetName);
        }

        // Create hidden sheets in the same batch
        const hiddenSheetObjects = [];
        for (const hidden of hiddenSheetsToAdd) {
          const sheet = sheets.add(hidden);
          hiddenSheetObjects.push({ sheet, name: hidden });
        }

        // Sync once after creating all sheets
        await context.sync();

        // Apply hidden property in batch
        for (const { sheet } of hiddenSheetObjects) {
          sheet.visibility = "Hidden";
        }

        // Create all tables sequentially instead of in parallel to avoid sync issues
        for (const sheetName of Object.keys(map)) {
          await createTable(context, sheetName as ModelType, intl);
        }

        // Create tables for all hidden sheets (recreate structure before population)
        for (const hidden of hiddenSheets) {
          try {
            await createHiddenTables(context, hidden);
          } catch (error) {
            // eslint-disable-next-line no-console
            console.error(`Error creating hidden table ${hidden}: ${error}`);
          }
        }
        await context.sync();

        // Add an explicit additional sync point to ensure all operations are complete
        await context.sync();

        // Add a longer delay here to ensure Excel has fully processed the table creation
        await new Promise((resolve) => setTimeout(resolve, 1000));

        // Populate dynamic hidden sheet (UnitsOfMeasure) before validations
        await populateUnitsOfMeasureHiddenSheet(context);

        // Add an explicit sync point here to ensure tables are fully populated
        await context.sync();

        // Add a small delay to ensure Excel has fully processed all data
        await new Promise((resolve) => setTimeout(resolve, 500));

        // Process one table at a time for validation setup
        for (const sheetName of Object.keys(map)) {
          try {
            await addLookupMultiSelectToLinkFieldsOnSheet(context, sheetName as ModelType);
            await context.sync(); // Sync after each sheet
          } catch (error) {
            console.error(`Error setting up validation for ${sheetName}: ${error}`);
          }
        }

        // Populate select lists after tables are created
        await populateSelectLists();
      } catch (error) {
        console.error("Error: " + error);
        throw error;
      }
    });
  } finally {
    disableSilentMode(); // Ensure silent mode is disabled after operations
  }
}

// Modify createSFFModuleSheetsAndTables to use try/catch for each step
export async function createSFFModuleSheetsAndTables(intl: IntlShape) {
  enableSilentMode();
  try {
    await Excel.run(async (context) => {
      try {
        // Clean up all existing event handlers first
        await cleanupAllEventHandlers(context);

        // Load worksheets once
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        // Get existing sheet names
        const existingSheetNames = new Set(sheets.items.map((sheet) => sheet.name));

        // Create needed sheets
        for (const name of Object.keys(map)) {
          if (!existingSheetNames.has(name)) {
            sheets.add(name);
          }
        }

        for (const name of Object.keys(mapSFFModel)) {
          if (!existingSheetNames.has(name)) {
            sheets.add(name);
          }
        }

        // Create hidden sheets one by one
        const hiddenSheetObjects = [];
        for (const name of hiddenSheets) {
          if (!existingSheetNames.has(name)) {
            const sheet = sheets.add(name);
            hiddenSheetObjects.push({ sheet, name });
          }
        }

        await context.sync();

        // Set hidden property
        for (const { sheet } of hiddenSheetObjects) {
          sheet.visibility = "Hidden";
        }

        // Process tables in a fixed sequence to avoid overlaps
        for (const sheetName of Object.keys(map)) {
          await createTable(context, sheetName as ModelType, intl);
        }

        for (const sheetName of Object.keys(mapSFFModel)) {
          await createTable(context, sheetName as SFFModelType, intl);
        }

        for (const hidden of hiddenSheets) {
          try {
            await createHiddenTables(context, hidden);
          } catch (error) {
            console.error(
              `Failed to create hidden table ${hidden}, continuing with others: ${error}`
            );
          }
        }

        // Set up fields
        for (const sheetName of Object.keys(map)) {
          try {
            await addLookupMultiSelectToLinkFieldsOnSheet(context, sheetName as ModelType);
          } catch (error) {
            console.error(`Error setting up fields for ${sheetName}: ${error}`);
          }
        }

        for (const sheetName of Object.keys(mapSFFModel)) {
          try {
            await addLookupMultiSelectToLinkFieldsOnSheet(context, sheetName as SFFModelType);
          } catch (error) {
            console.error(`Error setting up fields for ${sheetName}: ${error}`);
          }
        }

        // First populate hidden sheets with code lists (excluding UnitsOfMeasure which is handled separately)
        await populateHiddenSheets(context);
        await populateUnitsOfMeasureHiddenSheet(context);

        // Now populate standard code lists
        await populateCodeLists();
      } catch (error) {
        console.error(`Top-level error in createSFFModuleSheetsAndTables: ${error}`);
        throw error;
      }
    });
  } finally {
    disableSilentMode();

    // Re-register event handlers with a delay to ensure sync completes first
    setTimeout(async () => {
      try {
        await Excel.run(async () => {
          // Register all handlers including the ID column handlers
          await addLookupMultiSelectHandlerToAllTables();
        });
      } catch (error) {
        console.error("Failed to re-register event handlers:", error);
      }
    }, HANDLER_REGISTRATION_DELAY); // Add a bit extra time for safety
  }
}

async function addLookupMultiSelectToLinkFieldsOnSheet(
  context: Excel.RequestContext,
  sheetName: ModelType | SFFModelType
) {
  try {
    // Get all sheet/table info in one batch
    const sheet = context.workbook.worksheets.getItem(sheetName);
    context.trackedObjects.add(sheet); // Track the sheet

    const table = sheet.tables.getItem(sheetName);
    context.trackedObjects.add(table); // Track the table

    table.load("rows");
    await context.sync();

    const totalRows = table.rows;
    totalRows.load("count");
    sheet.load("name");
    await context.sync();

    const newClass = createInstance(sheetName);
    const fields = newClass.getAllFields();

    // Get all headers at once to reduce lookups
    const headerRange = table.getHeaderRowRange();
    context.trackedObjects.add(headerRange); // Track header range
    headerRange.load("values");
    await context.sync();
    const headers = headerRange.values[0];

    // Group fields by type
    const linkFields = fields.filter((field) => field.link);
    const selectFields = fields.filter((field) => field.type === "select");
    const multiselectFields = fields.filter((field) => field.type === "multiselect");

    // Process fields one at a time to avoid validation issues
    // Process all fields using a for...of loop instead of Promise.all to ensure sequential processing
    const allFields = [...linkFields, ...selectFields, ...multiselectFields];
    const validationResults = [];

    for (const field of allFields) {
      const fieldName = field.displayName || field.name;
      const columnIndex = headers.indexOf(fieldName);

      if (columnIndex === -1) {
        if (dialogHandler) {
          dialogHandler(
            { descriptor: { id: "generics.error" } },
            {
              descriptor: {
                id: "createTables.messages.error.fieldNotFound",
                defaultMessage: "Field {fieldName} not found in the sheet {sheetName}.",
              },
              values: { fieldName, sheetName: sheet.name },
            }
          );
        }
        continue; // Skip this field and continue with the next one
      }

      // Create range and track it immediately
      const range = sheet.getRangeByIndexes(1, columnIndex, totalRows.count, 1);
      context.trackedObjects.add(range); // Explicitly track the range
      range.load("address");
      await context.sync(); // Sync to ensure the range is loaded and available

      if (field.link) {
        // Set up validation for link fields
        await setupLinkFieldValidation(field, range, context);

        // Refresh range reference after validation setup
        // Notice we're keeping the same range object but refreshing its properties
        range.load("address");
        await context.sync();

        validationResults.push({ type: "link", field, range, address: range.address });
      } else if (field.type === "select" || field.type === "multiselect") {
        // Set up validation for select/multiselect fields
        await setupSelectFieldValidation(field, range, context);

        if (field.type === "multiselect") {
          // Refresh range reference after validation setup
          range.load("address");
          await context.sync();

          validationResults.push({ type: "multiselect", field, range, address: range.address });
        }
      }
    }

    // Register event handlers for all ranges
    for (const result of validationResults) {
      if (result.type === "link") {
        addLookupMultiSelectHandler(sheet, result.address);
      } else if (result.type === "multiselect") {
        addMultiSelectHandler(sheet, result.address);
      }
    }
  } catch (error) {
    console.error(`Error in addLookupMultiSelectToLinkFieldsOnSheet for ${sheetName}: ${error}`);
    throw error;
  }
}

// New helper function for setting up link field validation
async function setupLinkFieldValidation(
  field: any,
  range: Excel.Range,
  context: Excel.RequestContext
) {
  try {
    // We need to ensure the range is properly tracked and loaded before accessing its properties
    context.trackedObjects.add(range); // Explicitly track the object
    range.load("address"); // Load at minimum the address
    await context.sync(); // Sync to make sure data is available

    // Now set the validation rule
    // eslint-disable-next-line no-param-reassign
    range.dataValidation.rule = {
      list: {
        inCellDropDown: true,
        source: `=INDIRECT("${field.link.table.className}['@id]")`,
      },
    };

    // eslint-disable-next-line no-param-reassign
    range.dataValidation.errorAlert = {
      message: "Please select a value from the list.",
      showAlert: true,
      style: "Stop",
      title: "Invalid value",
    };

    await context.sync(); // Sync after setting validation to ensure it's applied
  } catch (error) {
    console.error(`Error in setupLinkFieldValidation for ${field.name}: ${error}`);
  }
}

// New helper function for setting up select/multiselect field validation
async function setupSelectFieldValidation(
  field: any,
  range: Excel.Range,
  context: Excel.RequestContext
) {
  try {
    // We need to ensure the range is properly tracked and loaded before accessing its properties
    context.trackedObjects.add(range); // Explicitly track the object
    range.load("address"); // Load at minimum the address
    await context.sync(); // Sync to make sure data is available

    // Find the appropriate source table for the dropdown
    let sourceName =
      [...predefinedCodeLists, ...hiddenSheets].find((listName) =>
        field.name.toLowerCase().includes(listName.toLowerCase())
      ) || field.name;

    // Special mapping for units of measure (field name/displayName does not match sheet name)
    if (
      field.name.toLowerCase().includes("unit_of_measure") ||
      (field.displayName && field.displayName.toLowerCase() === "unit_of_measure")
    ) {
      sourceName = "UnitsOfMeasure";
    }

    // Set validation rule
    // eslint-disable-next-line no-param-reassign
    range.dataValidation.rule = {
      list: {
        inCellDropDown: true,
        source: `=INDIRECT("${sourceName}[name]")`,
      },
    };

    // eslint-disable-next-line no-param-reassign
    range.dataValidation.errorAlert = {
      message: "Please select a value from the list.",
      showAlert: true,
      style: "Stop",
      title: "Invalid value",
    };

    await context.sync(); // Sync after setting validation to ensure it's applied
  } catch (error) {
    console.error(`Error in setupSelectFieldValidation for ${field.name}: ${error}`);
  }
}

// Ensure event handlers are checking for the correct conditions
function addLookupMultiSelectHandler(sheet: Excel.Worksheet, rangeAddress: string) {
  const handlerId = `${sheet.name}_lookup_${rangeAddress}`;

  // Remove any existing handler first
  if (registeredHandlers.has(handlerId)) {
    try {
      const handler = registeredHandlers.get(handlerId);
      if (handler) {
        handler.remove();
      }
    } catch (e) {
      console.error(`Failed to remove old handler ${handlerId}: ${e}`);
    }
    registeredHandlers.delete(handlerId);
  }

  // Add handler with improved range checking
  const handler = sheet.onChanged.add(async (e: Excel.WorksheetChangedEventArgs): Promise<void> => {
    // Skip if we're in silent mode or the event was triggered by this add-in
    if (inSilentMode || pendingOperations || e.triggerSource === "ThisLocalAddin") {
      return;
    }

    // Show warning if we don't have event details
    if (!e.details) {
      if (dialogHandler) {
        dialogHandler(
          { descriptor: { id: "generics.warning" } },
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

    // Check if this event is in our range - with extra logging
    const eventAddress = e.address;
    const isInRange = importedCellInRange(eventAddress, rangeAddress);

    // Only process the event if it's in our range
    if (isInRange) {
      await lookupMultiSelectEventHandler(e, rangeAddress);
    }
  });

  registeredHandlers.set(handlerId, handler);
}

// Apply the same pattern to the multiSelectHandler for consistency
function addMultiSelectHandler(sheet: Excel.Worksheet, rangeAddress: string) {
  const handlerId = `${sheet.name}_multiselect_${rangeAddress}`;

  // Remove any existing handler first
  if (registeredHandlers.has(handlerId)) {
    try {
      const handler = registeredHandlers.get(handlerId);
      if (handler) {
        handler.remove();
      }
    } catch (e) {
      console.error(`Failed to remove old handler ${handlerId}: ${e}`);
    }
    registeredHandlers.delete(handlerId);
  }

  // Add handler with improved range checking
  const handler = sheet.onChanged.add(async (e: Excel.WorksheetChangedEventArgs): Promise<void> => {
    // Skip if we're in silent mode or the event was triggered by this add-in
    if (inSilentMode || pendingOperations || e.triggerSource === "ThisLocalAddin") {
      return;
    }

    // Show warning if we don't have event details
    if (!e.details) {
      if (dialogHandler) {
        dialogHandler(
          { descriptor: { id: "generics.warning" } },
          {
            descriptor: {
              id: "eventHandler.messages.warning.changesForRangeNotSupported",
            },
          }
        );
      }
      return;
    }

    // Check if this event is in our range - with extra logging
    const eventAddress = e.address;
    const isInRange = importedCellInRange(eventAddress, rangeAddress);

    // Only process the event if it's in our range
    if (isInRange) {
      await multiSelectEventHandler(e, rangeAddress);
    }
  });

  registeredHandlers.set(handlerId, handler);
}

// Ensure the handler cleanup happens when appropriate
async function addLookupMultiSelectHandlerToAllTables() {
  await Excel.run(async (context) => {
    try {
      // Clean up existing handlers first to prevent duplicates
      await cleanupAllEventHandlers(context);

      // Register the ID column change handlers first
      await registerIdColumnChangeHandlers(context);

      // Now register the link field and multi-select field handlers
      await registerLinkAndMultiselectHandlers(context);
    } catch (error) {
      console.error("Error in addLookupMultiSelectHandlerToAllTables:", error);
    }
  });
}

// Make sure event handler is properly guarding against the event's address
async function lookupMultiSelectEventHandler(
  event: Excel.WorksheetChangedEventArgs,
  rangeAddress: string
) {
  // Skip if triggered by add-in or in silent mode
  if (event.triggerSource === "ThisLocalAddin" || inSilentMode || pendingOperations) {
    return;
  }

  // Check if the event is for a cell in our range - early exit if not
  if (!importedCellInRange(event.address, rangeAddress)) {
    return;
  }

  if (!event.details) {
    // Already handled in the caller
    return;
  }

  await Excel.run(async (context) => {
    try {
      // Load all necessary data at once
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load("name");
      const targetCell = event.getRange(context);
      targetCell.load("address, values, rowIndex, columnIndex");

      await context.sync();

      // Get table information
      const table = activeSheet.tables.getItem(activeSheet.name);
      const tableRange = table.getRange();
      const tableHeadersRange = table.getHeaderRowRange();
      tableHeadersRange.load("values");
      await context.sync();

      const headers = tableHeadersRange.values[0];
      const idColumnIndex = headers.indexOf("@id");
      const targetRowIndex = targetCell.rowIndex;
      const targetColumnIndex = targetCell.columnIndex;

      // Handle ID cell changes
      if (targetColumnIndex === idColumnIndex) {
        await handleIdCellChange(context, activeSheet, targetCell, tableRange, event.details);
        return;
      }

      // Exit if cell is not in the range we care about
      if (!importedCellInRange(targetCell.address, rangeAddress)) {
        return;
      }

      // Get ID cell and row data
      const idCell = activeSheet.getRangeByIndexes(targetRowIndex, idColumnIndex, 1, 1);
      idCell.load("values");
      await context.sync();

      // Get the field information
      const fieldName = headers[targetColumnIndex];
      if (!fieldName) {
        targetCell.values = [[event.details.valueBefore]];
        return;
      }

      // Handle linked field changes
      await handleLinkedFieldChange(
        context,
        activeSheet,
        targetCell,
        idCell,
        fieldName,
        event.details
      );
    } catch (error) {
      console.error("Error: " + error);
    }
  });
}

// New helper function to handle ID cell changes
async function handleIdCellChange(
  context: Excel.RequestContext,
  activeSheet: Excel.Worksheet,
  targetCell: Excel.Range,
  tableRange: Excel.Range,
  details: Excel.ChangedEventDetail
) {
  const newId = details.valueAfter.toString();
  const oldId = details.valueBefore.toString();
  const targetRowIndex = targetCell.rowIndex;

  // Load row data to check if it has content
  const row = tableRange.getRow(targetRowIndex);
  row.load("values");
  await context.sync();
  const rowValues = row.values[0];

  // Check if ID can be cleared (row must be empty)
  if (newId === "" || newId === null) {
    if (rowValues.some((v: any) => v !== "")) {
      // eslint-disable-next-line no-param-reassign
      targetCell.values = [[oldId]];
      showErrorDialog(
        "idCannotBeEmpty",
        "The id cannot be empty for a row with values. Please remove all values first."
      );
      return;
    }
    // Allow clearing ID of empty row
    // eslint-disable-next-line no-param-reassign
    targetCell.values = [[newId]];
    return;
  }

  // Check ID uniqueness
  const idColumn = tableRange.getColumn(targetCell.columnIndex);
  idColumn.load("values");
  await context.sync();

  if (idColumn.values.some((v: any[], i: number) => i !== targetRowIndex && v[0] === newId)) {
    // eslint-disable-next-line no-param-reassign
    targetCell.values = [[oldId]];
    showErrorDialog("newIdNotUnique", "The new id is not unique. Please enter a unique id.");
    return;
  }

  // Validate URL format
  try {
    new URL(newId);
  } catch (error) {
    // eslint-disable-next-line no-param-reassign
    targetCell.values = [[oldId]];
    showErrorDialog("newIdNotValidURL", "The new id is not a valid URL. Please enter a valid URL.");
    return;
  }

  // ID is valid, update it
  // eslint-disable-next-line no-param-reassign
  targetCell.values = [[newId]];

  // Update related fields in other tables
  const sheetName = activeSheet.name as ModelType | SFFModelType;
  const model = createInstance(sheetName);

  // Get all linked fields that need updating
  const linkedFields = model
    .getAllFields()
    .filter(
      (field) =>
        field.link &&
        (!(ignoredFields as Record<string, any>)[field.link.table.className] ||
          !(ignoredFields as Record<string, any>)[field.link.table.className].includes(
            field.link.field
          ))
    );

  // Process updates sequentially for reliability (like original implementation)
  for (const field of linkedFields) {
    try {
      await updateRelatedFieldsValues(
        context,
        field.link?.table.className,
        field.link?.field || "",
        oldId,
        newId
      );
    } catch (error) {
      console.error(`Error updating ${field.link?.table.className}.${field.link?.field}: ${error}`);
    }
  }
}

// New helper function to handle linked field changes
async function handleLinkedFieldChange(
  context: Excel.RequestContext,
  activeSheet: Excel.Worksheet,
  targetCell: Excel.Range,
  idCell: Excel.Range,
  fieldName: string,
  details: Excel.ChangedEventDetail
) {
  const newValue: string = details.valueAfter.toString();
  const oldValue: string = details.valueBefore.toString();

  // Quick validation - If ID is empty, prevent changes immediately
  if (!idCell.values || !idCell.values[0][0]) {
    // eslint-disable-next-line no-param-reassign
    targetCell.values = [[oldValue]];
    return;
  }

  // Process the multi-select style update early
  const processedValue = processMultiSelectValue(newValue, oldValue);
  if (processedValue === null) {
    // No change needed
    return;
  }

  // Apply the new value IMMEDIATELY for fast feedback
  // eslint-disable-next-line no-param-reassign
  targetCell.values = [[processedValue]];

  // Sync immediately to show the user their change
  await context.sync();

  // Now perform validation and related operations in the background
  try {
    // Get field definition
    const model = createInstance(activeSheet.name as ModelType | SFFModelType);
    const field = model.getFieldByName(fieldName);

    if (!field || !field.link) {
      throw new Error("Field not found");
    }

    // Validate the new value exists in the related sheet
    if (
      !(await checkIfAllValuesExistInRelatedSheet(context, field.link.table.className, newValue))
    ) {
      throw new Error("Invalid value");
    }

    // Update bi-directional links in related tables if needed
    if (
      field.link.table.className !== activeSheet.name &&
      (!(ignoredFields as Record<string, any>)[field.link.table.className] ||
        !(ignoredFields as Record<string, any>)[field.link.table.className].includes(
          field.link.field
        ))
    ) {
      // Use a small timeout to ensure UI updates have been processed
      await new Promise((resolve) => setTimeout(resolve, 50));

      // Update related fields in background
      await handleLinkedFieldsInRelatedSheet(
        context,
        field.link.table.className,
        field.link.field,
        idCell.values[0][0].toString(),
        processedValue
      );
    }
  } catch (error: any) {
    // If validation fails, revert the cell value
    if (error.message === "Field not found") {
      // eslint-disable-next-line no-param-reassign
      targetCell.values = [[oldValue]];
      showErrorDialog("fieldNotFound", "Field not found.");
    } else if (error.message === "Invalid value") {
      // eslint-disable-next-line no-param-reassign
      targetCell.values = [[oldValue]];
      showErrorDialog("invalidValue", "Invalid value. Please select a value from the list.");
    } else {
      console.error("Error in handleLinkedFieldChange:", error);
      // eslint-disable-next-line no-param-reassign
      targetCell.values = [[oldValue]];
    }
    await context.sync(); // Make sure the reversion is visible
  }
}

// Helper function to show error dialogs
function showErrorDialog(errorId: string, defaultMessage: string) {
  if (dialogHandler) {
    showDialog(dialogHandler, "error", errorId, defaultMessage);
  }
}

// These functions need significant optimization
export async function populateSelectLists() {
  enableSilentMode();
  try {
    await Excel.run(async (context) => {
      try {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        // Check which sheets actually exist
        const existingSheetNames = new Set(sheets.items.map((sheet) => sheet.name));
        const availableSheets = hiddenSheets.filter((name) => existingSheetNames.has(name));

        // Load all code list data in parallel with enhanced error handling
        const codeListData = await Promise.all(
          availableSheets.map(async (codeList) => {
            try {
              const data = await debugCodeListData(codeList);
              return {
                name: codeList,
                data,
                error: null,
              };
            } catch (error) {
              console.error(`Failed to fetch data for ${codeList}: ${error}`);
              return {
                name: codeList,
                data: null,
                error,
              };
            }
          })
        );

        // Process each code list with batch operations and better error handling
        for (const { name: codeList, data, error } of codeListData) {
          if (error) {
            continue;
          }

          if (!data || data.length === 0) {
            continue;
          }

          try {
            const sheet = sheets.getItem(codeList);

            // Verify sheet has tables
            sheet.load("tables/items/length");
            await context.sync();

            if (sheet.tables.items.length === 0) {
              await createHiddenTables(context, codeList);
            }

            const table = sheet.tables.getItem(codeList);
            const tableRange = table.getRange();

            // Populate with optimized function
            await populateCodeListBatched(context, codeList, tableRange, data);
          } catch (error) {
            console.error(`Error processing ${codeList}: ${error}`);
          }
        }
      } catch (error) {
        console.error(`Error in populateSelectLists: ${error}`);
      }
    });
  } finally {
    disableSilentMode();
  }
}

export async function populateCodeLists() {
  enableSilentMode();
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      // Check which predefined lists actually exist as sheets
      const existingSheetNames = new Set(sheets.items.map((sheet) => sheet.name));
      const availableLists = predefinedCodeLists.filter((name) => existingSheetNames.has(name));

      // Load all code list data in parallel with better error handling
      const codeListData = await Promise.all(
        availableLists.map(async (codeList) => {
          try {
            const data = await debugCodeListData(codeList);
            return {
              name: codeList,
              data,
              error: null,
            };
          } catch (error) {
            return {
              name: codeList,
              data: null,
              error,
            };
          }
        })
      );

      // Process each code list with batch operations
      for (const { name: codeList, data, error } of codeListData) {
        if (error) {
          continue;
        }

        if (!data || data.length === 0) {
          continue;
        }

        try {
          const sheet = sheets.getItem(codeList);
          const table = sheet.tables.getItem(codeList);

          // Get table range and data in one batch
          const tableRange = table.getRange();
          tableRange.load("values");
          await context.sync();

          // Populate with optimized function
          await populateCodeListBatched(context, codeList, tableRange, data);
        } catch (error) {
          console.error(`Error populating ${codeList}: ${error}`);
        }
      }
    });
  } finally {
    disableSilentMode();
  }
}

// Fix the populateCodeListBatched function to properly load values
async function populateCodeListBatched(
  context: Excel.RequestContext,
  tableName: string,
  tableRange: Excel.Range,
  data: any[]
) {
  try {
    // Check if we actually have data to populate
    if (!data || data.length === 0) {
      return;
    }

    // Ensure table range values are loaded
    tableRange.load("values");
    await context.sync();

    // Now we can safely access values
    const tableValues = tableRange.values;

    // Determine column structure
    const headers = hiddenSheets.includes(tableName) ? ["id", "name"] : Object.keys(data[0]);
    const idColumnName = hiddenSheets.includes(tableName) ? "id" : "@id";
    const idColumnIndex = tableValues[0].indexOf(idColumnName);

    if (idColumnIndex === -1) {
      return;
    }

    // Find first empty row - more robust approach for new tables
    let firstEmptyRowIndex = 1; // Start at first data row (after header)

    // Extract existing IDs for de-duplication
    const idToRowIndex = new Map<string, number>();

    // Skip completely empty tables (only header row)
    if (tableValues.length > 1) {
      for (let i = 1; i < tableValues.length; i++) {
        const id = tableValues[i][idColumnIndex]?.toString();
        if (id) {
          idToRowIndex.set(id, i);
        } else if (firstEmptyRowIndex === 1) {
          firstEmptyRowIndex = i;
        }
      }
    }

    // Process data in batches to avoid overloading Excel
    const MAX_BATCH_SIZE = 50;
    const BATCH_COUNT = Math.ceil(data.length / MAX_BATCH_SIZE);

    for (let batchIndex = 0; batchIndex < BATCH_COUNT; batchIndex++) {
      const batchStart = batchIndex * MAX_BATCH_SIZE;
      const batchEnd = Math.min(batchStart + MAX_BATCH_SIZE, data.length);
      const batchData = data.slice(batchStart, batchEnd);

      // Map header names to column indices for each batch
      const headerIndices = new Map<string, number>();
      headers.forEach((header) => {
        const index = tableValues[0].indexOf(header);
        if (index !== -1) {
          headerIndices.set(header, index);
        }
      });

      // Prepare updates for this batch
      for (const item of batchData) {
        const itemId = item["@id"]; // Use @id consistently

        if (!itemId) {
          continue;
        }

        // Get or create row index for this item
        let rowIndex = idToRowIndex.get(itemId);
        if (rowIndex === undefined) {
          rowIndex = firstEmptyRowIndex++;

          // Add ID to the row
          const idCell = tableRange.getCell(rowIndex, idColumnIndex);
          idCell.values = [[itemId]];
        }

        // Update all other fields
        for (const header of headers) {
          const colIndex = headerIndices.get(header);
          if (colIndex === undefined) continue;

          // Map property names based on table type
          let value;

          if (hiddenSheets.includes(tableName)) {
            // Hidden sheets use id/name format
            if (header === "id") {
              value = item["@id"];
            } else if (header === "name") {
              value = item["hasName"] || item["name"] || item["prefLabel"] || "";
            }
          } else {
            // Standard code lists
            value = item[header] || "";
          }

          // Only update if we have a value
          if (value !== undefined && colIndex !== idColumnIndex) {
            const cell = tableRange.getCell(rowIndex, colIndex);
            cell.values = [[value]];
          }
        }
      }

      // Sync after each batch
      await context.sync();
    }
  } catch (error) {
    console.error(`Error populating ${tableName}: ${error}`);
  }
}

// Let's also add a helper to verify API data is coming through correctly
async function debugCodeListData(tableName: string) {
  try {
    let data;
    try {
      data = await getCodeListByTableName(tableName);
    } catch (apiError) {
      console.error(`API ERROR for ${tableName}: ${apiError}`);
      return null;
    }

    if (!data) {
      return null;
    }

    if (!Array.isArray(data)) {
      return null;
    }

    if (data.length > 0) {
      try {
        // Safely log first item (handle circular references)
        const sampleItem = { ...data[0] };
        const safeItem: Record<string, unknown> = {};

        // Extract just the key properties for logging
        for (const key of ["@id", "hasName", "name", "prefLabel"]) {
          if ((sampleItem as Record<string, unknown>)[key]) {
            safeItem[key] = (sampleItem as Record<string, unknown>)[key];
          }
        }
      } catch (err) {
        console.error(`Error logging sample for ${tableName}:`, err);
      }
    }

    return data;
  } catch (error) {
    console.error(`Error fetching data for ${tableName}: ${error}`);
    return null;
  }
}

// Improved event handler for multi-select fields
async function multiSelectEventHandler(
  event: Excel.WorksheetChangedEventArgs,
  rangeAddress: string
) {
  // Skip if triggered by add-in, in silent mode, or doesn't have details
  if (event.triggerSource === "ThisLocalAddin" || inSilentMode || !event.details) {
    if (!event.details && dialogHandler && !inSilentMode) {
      // Only show warning if not in silent mode
      dialogHandler(
        { descriptor: { id: "generics.warning" } },
        {
          descriptor: {
            id: "eventHandler.messages.warning.changesForRangeNotSupported",
          },
        }
      );
    }
    return;
  }

  // Get the values immediately for local processing
  const newValue = event.details.valueAfter.toString();
  const oldValue = event.details.valueBefore.toString();

  // Handle common case - empty to empty
  if ((!newValue || newValue === "") && (!oldValue || oldValue === "")) {
    return; // No change needed
  }

  // Process the change locally first to minimize Excel API calls
  const processedValue = processMultiSelectValue(newValue, oldValue);
  if (processedValue === null) {
    return; // No change needed
  }

  // Create a unique key for this cell to handle debouncing
  const cellKey = `${event.worksheetId}_${event.address}`;

  // Apply changes with two-phase approach:
  // 1. First immediate feedback
  // 2. Then proper formatting with debounce
  await Excel.run(async (context) => {
    try {
      const targetCell = event.getRange(context);
      targetCell.load("address");
      await context.sync();

      // Exit early if not in our range
      if (!importedCellInRange(targetCell.address, rangeAddress)) {
        return;
      }

      console.log("Processing multi-select change:", newValue, oldValue, processedValue);

      // IMMEDIATE FEEDBACK: Show the new value right away with formatting
      // to indicate processing is happening
      targetCell.values = [[processedValue]];
      targetCell.format.font.italic = true;
      targetCell.format.fill.color = "#F5F5F5"; // Light gray background
      await context.sync();

      // Debounce the final formatting to avoid rapid successive updates
      debounce(
        cellKey,
        async () => {
          await Excel.run(async (innerContext) => {
            try {
              const finalCell = innerContext.workbook.worksheets
                .getActiveWorksheet()
                .getRange(targetCell.address);

              // Apply final value with normal formatting
              finalCell.values = [[processedValue]];
              finalCell.format.font.italic = false;
              finalCell.format.fill.clear();
              await innerContext.sync();
            } catch (error) {
              console.error("Error in multiSelectEventHandler debounced final formatting:", error);
            }
          });
        },
        300
      ); // 300ms debounce - adjust as needed for best UX
    } catch (error) {
      console.error("Error in multiSelectEventHandler immediate feedback:", error);
    }
  });
}

// Split the code list population into a separate function specifically for hidden sheets
async function populateHiddenSheets(context: Excel.RequestContext) {
  try {
    // Check which sheets actually exist
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const existingSheetNames = new Set(sheets.items.map((sheet) => sheet.name));
    // Exclude UnitsOfMeasure here; it has a dedicated population function
    const availableSheets = hiddenSheets
      .filter((name) => name !== "UnitsOfMeasure")
      .filter((name) => existingSheetNames.has(name));

    // Process each sheet's code list one by one
    for (const codeList of availableSheets) {
      try {
        const data = await getCodeListByTableName(codeList);

        if (!data || data.length === 0) {
          continue;
        }

        const sheet = sheets.getItem(codeList);

        // Force recreate table to avoid conflicts
        try {
          sheet.tables.getItem(codeList).delete();
        } catch (e) {
          console.warn(`No existing table to delete for ${codeList}`);
        }
        await context.sync();

        const table = await createHiddenTables(context, codeList);

        // Populate data
        const tableRange = table.getRange();
        await populateCodeListBatchedSimple(context, tableRange, data);
      } catch (error) {
        console.error(`Error processing ${codeList}: ${error}`);
      }
    }
  } catch (error) {
    console.error(`Error in populateHiddenSheets: ${error}`);
  }
}

// Simplified version of populateCodeListBatched specifically for hidden sheets
async function populateCodeListBatchedSimple(
  context: Excel.RequestContext,
  tableRange: Excel.Range,
  data: any[]
) {
  try {
    if (!data || data.length === 0) {
      return;
    }

    // Load the table range dimensions first
    tableRange.load("rowCount,columnCount,worksheet");
    await context.sync();

    // Create data matrix with header and all data rows
    const values = [];

    // Get header row properly - don't duplicate it
    values.push(["id", "name"]);

    // Add data rows
    for (const item of data) {
      values.push([item["@id"] || "", item["hasName"] || item["name"] || item["prefLabel"] || ""]);
    }

    // Either create a fresh range or use a different approach
    try {
      // Method 1: Create a completely fresh range with exact size needed
      // This avoids modifying the parameter directly
      const worksheet = tableRange.worksheet;
      const freshRange = worksheet.getRangeByIndexes(0, 0, values.length, 2); // start at A1, exact size
      freshRange.values = values;
      await context.sync();
    } catch (rangeError) {
      // Method 2: Use a new range reference instead of modifying parameter directly
      try {
        // Get worksheet reference
        const worksheet = tableRange.worksheet;

        // Create a fresh range covering the same area
        const targetRange = worksheet.getRange(tableRange.address);
        // Clear the range without modifying parameter
        targetRange.clear();
        await context.sync();

        // Set header row using the worksheet instead of the parameter
        worksheet.getRange("A1:B1").values = [["id", "name"]];

        // Set data rows one by one using new range references
        for (let i = 0; i < data.length; i++) {
          const item = data[i];
          if (i + 1 < tableRange.rowCount) {
            // Make sure we don't exceed range
            const cellAddress = `A${i + 2}:B${i + 2}`;
            worksheet.getRange(cellAddress).values = [
              [item["@id"] || "", item["hasName"] || item["name"] || item["prefLabel"] || ""],
            ];
          }
        }
        await context.sync();
      } catch (alternateError) {
        console.warn(`Alternate method also failed: ${alternateError}`);
      }
    }
  } catch (error) {
    console.error(`Error in populateCodeListBatchedSimple: ${error}`);
  }
}

// Add a specific function to register ID column change handlers
async function registerIdColumnChangeHandlers(context: Excel.RequestContext) {
  try {
    // Get all sheet names from both maps
    const sheetNames = [...Object.keys(map), ...Object.keys(mapSFFModel)];

    // Get all relevant sheets
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const relevantSheets = sheets.items.filter((sheet) => sheetNames.includes(sheet.name));

    for (const sheet of relevantSheets) {
      // Don't register handlers for hidden sheets
      if (hiddenSheets.includes(sheet.name)) {
        continue;
      }

      try {
        // Find the ID column in each table
        const table = sheet.tables.getItem(sheet.name);
        const headerRange = table.getHeaderRowRange();
        headerRange.load("values");
        await context.sync();

        const headers = headerRange.values[0];
        const idColumnIndex = headers.indexOf("@id");

        if (idColumnIndex === -1) {
          continue;
        }

        // Get row count
        table.load("rows/count");
        await context.sync();

        // Register handler for the ID column specifically
        const idColumnRange = sheet.getRangeByIndexes(1, idColumnIndex, table.rows.count, 1);
        idColumnRange.load("address");
        await context.sync();

        // Register a specific handler just for the ID column
        addIdChangeHandler(sheet, idColumnRange.address);
      } catch (error) {
        console.error(`Error setting up ID column handler for sheet ${sheet.name}: ${error}`);
      }
    }
  } catch (error) {
    console.error("Error in registerIdColumnChangeHandlers:", error);
  }
}

// Add a specific handler function for ID columns
function addIdChangeHandler(sheet: Excel.Worksheet, rangeAddress: string) {
  const handlerId = `${sheet.name}_idchange_${rangeAddress}`;

  // Remove any existing handler first
  if (registeredHandlers.has(handlerId)) {
    try {
      const handler = registeredHandlers.get(handlerId);
      if (handler) {
        handler.remove();
      }
    } catch (e) {
      console.warn(`Failed to remove old ID change handler ${handlerId}: ${e}`);
    }
    registeredHandlers.delete(handlerId);
  }

  // Add handler specific for ID column changes - with enhanced logging
  const handler = sheet.onChanged.add(async (e: Excel.WorksheetChangedEventArgs): Promise<void> => {
    // Ensure we have proper event details
    if (!e.details) {
      return;
    }

    // Check silent mode with clear log
    if (inSilentMode || pendingOperations) {
      return;
    }

    if (e.triggerSource === "ThisLocalAddin") {
      return;
    }

    // Special handling for ID column changes - more robust
    // This normalizes the addresses to a standard format
    const normalizedCellAddress = normalizeAddress(e.address);
    const cellColumn = normalizedCellAddress.match(/([A-Z]+)(\d+)/)?.[1];
    const cellRow = parseInt(normalizedCellAddress.match(/([A-Z]+)(\d+)/)?.[2] || "0", 10);

    // Check if this is column A (the ID column) and within the row range
    // We explicitly check for column A since that's where IDs are
    const isIdCell = cellColumn === "A" && cellRow >= 2;

    if (!isIdCell) {
      return;
    }
    await idChangeEventHandler(e, sheet);
  });

  registeredHandlers.set(handlerId, handler);
}

// Add a helper function to normalize cell addresses
function normalizeAddress(address: string): string {
  // If address already has sheet prefix, extract just the cell part
  if (address.includes("!")) {
    return address.split("!")[1];
  }

  // Otherwise return the address as-is
  return address;
}

// Update the idChangeEventHandler to provide better logging and more reliable detection
async function idChangeEventHandler(
  event: Excel.WorksheetChangedEventArgs,
  sheet: Excel.Worksheet
) {
  await Excel.run(async (context) => {
    try {
      // Extract values from the event
      const newId = event.details?.valueAfter.toString() || "";
      const oldId = event.details?.valueBefore.toString() || "";

      // Skip if no change
      if (newId === oldId) {
        return;
      }

      // Get cell and validate
      const targetCell = event.getRange(context);
      targetCell.load("rowIndex");
      sheet.load("name");
      await context.sync();

      // Get table and validate the ID change
      const table = context.workbook.worksheets.getItem(sheet.name).tables.getItem(sheet.name);
      const tableRange = table.getRange();

      // Fix: Load tableRange properties before accessing them
      tableRange.load("rowIndex");
      await context.sync();

      // Load row data to check if it has content
      const row = tableRange.getRow(targetCell.rowIndex - tableRange.rowIndex);
      row.load("values");
      await context.sync();

      // Validation checks (similar to handleIdCellChange)
      if (newId === "") {
        if (row.values[0].some((v: any) => v !== "")) {
          targetCell.values = [[oldId]];
          showErrorDialog("idCannotBeEmpty", "The id cannot be empty for a row with values");
          return;
        }
      } else {
        // Force these updates to run in silent mode to avoid triggering more events
        enableSilentMode();
        try {
          // Find all tables that have fields linking to this table
          for (const modelName of [...Object.keys(map), ...Object.keys(mapSFFModel)]) {
            if (modelName === sheet.name) continue;

            try {
              const linkedModel = createInstance(modelName as ModelType | SFFModelType);
              const linkedFields = linkedModel
                .getAllFields()
                .filter(
                  (field) =>
                    field.link && field.link.table && field.link.table.className === sheet.name
                );

              if (linkedFields.length > 0) {
                for (const field of linkedFields) {
                  try {
                    if (field.link && field.link.field) {
                      const linkedFieldName = field.displayName || field.name;

                      // Use direct updates to Excel rather than the helper function
                      await directlyUpdateLinkedFieldValues(
                        context,
                        modelName,
                        linkedFieldName,
                        oldId,
                        newId
                      );
                    }
                  } catch (err) {
                    console.error(`Error updating field ${field.name}: ${err}`);
                  }
                }
              }
            } catch (err) {
              console.error(`Error checking model ${modelName}: ${err}`);
            }
          }
        } finally {
          // Always disable silent mode when done
          disableSilentMode();
        }
      }
    } catch (error) {
      console.error(`Error in idChangeEventHandler: ${error}`);
    }
  });
}

// Add a new direct update function that works with less intermediate steps
async function directlyUpdateLinkedFieldValues(
  context: Excel.RequestContext,
  tableName: string,
  fieldName: string,
  oldValue: string,
  newValue: string
) {
  try {
    // Get worksheet
    const sheet = context.workbook.worksheets.getItem(tableName);
    const table = sheet.tables.getItem(tableName);

    // Load headers and find target column
    const headers = table.getHeaderRowRange();
    headers.load("values");
    await context.sync();

    const headerValues = headers.values[0];
    const columnIndex = headerValues.indexOf(fieldName);

    if (columnIndex === -1) {
      return;
    }

    // Get all data at once
    const dataRange = table.getDataBodyRange();
    dataRange.load("values");
    dataRange.load("rowCount");
    await context.sync();

    // Store values locally for efficient processing
    const rowCount = dataRange.rowCount;
    const allValues = dataRange.values;
    let updatesNeeded = false;

    // Process each row
    for (let i = 0; i < rowCount; i++) {
      const cellValue = allValues[i][columnIndex];

      if (cellValue === oldValue) {
        // Exact match
        dataRange.getCell(i, columnIndex).values = [[newValue]];
        updatesNeeded = true;
      } else if (cellValue && typeof cellValue === "string" && cellValue.includes(oldValue)) {
        // Check for multi-select list
        const valueItems = cellValue.split(", ");
        if (valueItems.includes(oldValue)) {
          const updatedItems = valueItems.map((item) => (item === oldValue ? newValue : item));
          dataRange.getCell(i, columnIndex).values = [[updatedItems.join(", ")]];
          updatesNeeded = true;
        }
      }
    }

    if (updatesNeeded) {
      await context.sync();
    }
  } catch (error) {
    console.error(`Error in directlyUpdateLinkedFieldValues: ${error}`);
  }
}

// Add a new function to specifically register handlers for link fields and multi-select fields
async function registerLinkAndMultiselectHandlers(context: Excel.RequestContext) {
  try {
    // Get all sheet names from both maps
    const sheetNames = [...Object.keys(map), ...Object.keys(mapSFFModel)];

    // Get all sheets
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const relevantSheets = sheets.items.filter(
      (sheet) => sheetNames.includes(sheet.name) && !hiddenSheets.includes(sheet.name)
    );

    for (const sheet of relevantSheets) {
      try {
        // Get the model and fields
        const sheetModel = createInstance(sheet.name as ModelType | SFFModelType);
        const allFields = sheetModel.getAllFields();

        // Find link fields and multi-select fields
        const linkFields = allFields.filter((field) => field.link);
        const multiselectFields = allFields.filter((field) => field.type === "multiselect");

        if (linkFields.length === 0 && multiselectFields.length === 0) {
          continue; // Skip if no relevant fields
        }

        // Get table information
        const table = sheet.tables.getItem(sheet.name);
        const headerRange = table.getHeaderRowRange();
        headerRange.load("values");
        table.load("rows/count");
        await context.sync();

        const headers = headerRange.values[0];
        const rowCount = table.rows.count;

        // Process link fields
        for (const field of linkFields) {
          const fieldName = field.displayName || field.name;
          const columnIndex = headers.indexOf(fieldName);

          if (columnIndex === -1) {
            continue;
          }

          // Create range and register handler
          const range = sheet.getRangeByIndexes(1, columnIndex, rowCount, 1);
          range.load("address");
          await context.sync();

          // Add explicit debug logging to check range addresses
          addLookupMultiSelectHandler(sheet, range.address);
        }

        // Process multi-select fields
        for (const field of multiselectFields) {
          const fieldName = field.displayName || field.name;
          const columnIndex = headers.indexOf(fieldName);

          if (columnIndex === -1) {
            continue;
          }

          // Create range and register handler
          const range = sheet.getRangeByIndexes(1, columnIndex, rowCount, 1);
          range.load("address");
          await context.sync();

          // Add explicit debug logging to check range addresses
          addMultiSelectHandler(sheet, range.address);
        }
      } catch (error) {
        console.error(`Error setting up handlers for ${sheet.name}: ${error}`);
      }
    }
  } catch (error) {
    console.error("Error in registerLinkAndMultiselectHandlers:", error);
  }
}

// Dedicated population for UnitsOfMeasure since it comes from a different source and already filtered
async function populateUnitsOfMeasureHiddenSheet(context: Excel.RequestContext) {
  const sheetName = "UnitsOfMeasure";
  try {
    // Ensure sheet exists
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    if (!sheets.items.some((s) => s.name === sheetName)) return;

    // Recreate the table to ensure clean state
    try {
      const sheet = sheets.getItem(sheetName);
      sheet.load("tables/items/length");
      await context.sync();
      if (sheet.tables.items.length > 0) {
        try {
          sheet.tables.getItem(sheetName).delete();
          await context.sync();
        } catch (_) {
          /* ignore */
        }
      }
    } catch (_) {
      /* ignore */
    }

    // Create empty table structure (id, name)
    await createHiddenTables(context, sheetName);
    await context.sync();

    // Fetch units
    const units = await getUnitOptions();
    if (!units || units.length === 0) return;

    // Populate
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const table = sheet.tables.getItem(sheetName);
    const tableRange = table.getRange();
    const values = [["id", "name"], ...units.map((u) => [u.id, u.name])];
    const targetRange = sheet.getRangeByIndexes(0, 0, values.length, 2);
    targetRange.values = values;
    await context.sync();
  } catch (error) {
    // eslint-disable-next-line no-console
    console.error(`Error populating ${sheetName}: ${error}`);
  }
}
