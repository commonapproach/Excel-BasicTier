/* global console, Excel */

export async function createHiddenTables(context: Excel.RequestContext, tableName: string) {
  try {
    const sheet = context.workbook.worksheets.getItem(tableName);
    const headers = ["id", "name"];

    // Clear the entire sheet first to ensure no tables exist
    const usedRange = sheet.getUsedRange();
    usedRange.clear();
    await context.sync();

    // Delete all existing tables on the sheet
    sheet.load("tables/items/length");
    sheet.load("tables/items");
    await context.sync();

    if (sheet.tables.items.length > 0) {
      // Load table names for better logging
      for (const tbl of sheet.tables.items) {
        tbl.load("name");
      }
      await context.sync();

      // Log and delete each table
      for (const tbl of sheet.tables.items) {
        tbl.delete();
      }
      await context.sync();
    }

    // Now the sheet should be clear, create the table
    const range = sheet.getRange("A1:B1");
    range.values = [headers];
    await context.sync();

    // Create table now that we've ensured no conflicts
    const table = sheet.tables.add(range.getResizedRange(1000, 0), true);
    table.name = tableName;
    table.showTotals = false;

    // Format the table
    const tableRange = table.getRange();
    tableRange.format.wrapText = true;
    tableRange.format.verticalAlignment = "Center";
    table.getHeaderRowRange().format.columnWidth = 220;
    tableRange.format.autofitColumns();
    tableRange.format.autofitRows();

    await context.sync();
    return table;
  } catch (error) {
    console.error(`Error in createHiddenTables for ${tableName}: ${error}`);
    throw error;
  }
}
