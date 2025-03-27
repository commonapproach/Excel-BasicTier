/**
 * Excel utility functions for working with ranges, cell values, and other common operations
 */

/**
 * Checks if a cell address is within a specified range
 * @param cellAddress - The address of the cell to check (e.g., "Sheet1!A1")
 * @param rangeAddress - The address of the range to check against (e.g., "A1:B10")
 * @returns True if the cell is within the range, false otherwise
 */
export function isCellInRange(cellAddress: string, rangeAddress: string): boolean {
  if (!cellAddress || !rangeAddress) return false;

  // Extract cell address from sheet reference if present
  let targetCellAddress = cellAddress;
  if (cellAddress.includes("!")) {
    targetCellAddress = cellAddress.split("!")[1];
  }

  // Extract range parts (might include sheet name)
  const rangeParts = rangeAddress.split(":");
  if (rangeParts.length !== 2) return false;

  // Handle sheet name in range if present
  let rangeStart = rangeParts[0];
  let rangeEnd = rangeParts[1];

  if (rangeStart.includes("!")) {
    const parts = rangeStart.split("!");
    rangeStart = parts[1];
    rangeEnd = rangeEnd.includes("!") ? rangeEnd.split("!")[1] : rangeEnd;
  }

  // Using regex patterns for reliable matching
  const startCellMatch = rangeStart.match(/([A-Z]+)(\d+)/);
  const endCellMatch = rangeEnd.match(/([A-Z]+)(\d+)/);
  const cellMatch = targetCellAddress.match(/([A-Z]+)(\d+)/);

  if (!startCellMatch || !endCellMatch || !cellMatch) {
    return false;
  }

  // Convert column letters to numbers for easier comparison
  const startColNum = columnLetterToNumber(startCellMatch[1]);
  const endColNum = columnLetterToNumber(endCellMatch[1]);
  const cellColNum = columnLetterToNumber(cellMatch[1]);

  const startRow = parseInt(startCellMatch[2], 10);
  const endRow = parseInt(endCellMatch[2], 10);
  const cellRow = parseInt(cellMatch[2], 10);

  return (
    cellColNum >= startColNum && cellColNum <= endColNum && cellRow >= startRow && cellRow <= endRow
  );
}

/**
 * Converts Excel column letter to number (e.g., A -> 1, Z -> 26, AA -> 27)
 * @param column - The column letter(s) to convert
 * @returns The column number
 */
export function columnLetterToNumber(column: string): number {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Compares two arrays for equality
 * @param a - First array
 * @param b - Second array
 * @returns True if arrays have same length and elements, false otherwise
 */
export function arraysEqual(a: any[], b: any[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}

/**
 * Processes multi-select values for adding/removing items
 * @param newValue - The new value being set
 * @param oldValue - The current value in the cell
 * @returns The processed value or null if no changes are needed
 */
export function processMultiSelectValue(newValue: string, oldValue: string): string | null {
  // Handle edge cases
  if (newValue === oldValue) return null;

  // Use sets for more efficient operations with large lists
  const oldValuesSet = new Set(oldValue ? oldValue.split(", ").filter(Boolean) : []);

  // Handle common single-value case efficiently
  if (!newValue) {
    // Clearing the value
    return "";
  } else if (!oldValue) {
    // First value being added
    return newValue;
  } else if (!oldValuesSet.has(newValue)) {
    // Add the new value (not present in old values)
    oldValuesSet.add(newValue);
    return Array.from(oldValuesSet).join(", ");
  } else {
    // Toggle behavior - remove the value
    oldValuesSet.delete(newValue);
    return Array.from(oldValuesSet).join(", ");
  }
}

/**
 * Shows a dialog with error or warning message
 * @param dialogHandler - The dialog handler function to use
 * @param errorId - The ID of the error message
 * @param defaultMessage - The default message to show
 */
export function showDialog(
  dialogHandler: any,
  type: "error" | "warning",
  errorId: string,
  defaultMessage: string
): void {
  if (!dialogHandler) return;

  dialogHandler(
    { descriptor: { id: `generics.${type}` } },
    {
      descriptor: {
        id: `eventHandler.messages.${type}.${errorId}`,
        defaultMessage,
      },
    }
  );
}
