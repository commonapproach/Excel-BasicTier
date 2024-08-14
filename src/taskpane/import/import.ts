import { TableInterface } from '../domain/interfaces/table.interface';
import { IndicatorReport, ModelType, Organization, map } from '../domain/models';
import { validate } from '../domain/validation/validator';
import { createSheetsAndTables } from '../taskpane';

/* global Excel console */
export async function importData(
  jsonData: any,
  setDialogContent: (header: string, content: string, nextCallBack?: Function) => void,
  setIsImporting: (value: boolean) => void
) {
  await Excel.run(async (context) => {
    if (validateIfEmptyFile(jsonData)) {
      setDialogContent('Error!', 'Table data is empty or not an array');
      setIsImporting(false);
      return;
    }

    if (!doAllRecordsHaveId(jsonData)) {
      setDialogContent('Error!', 'All records must have an <b>@id</b> property.');
      setIsImporting(false);
      return;
    }

    // Remove duplicated links
    // eslint-disable-next-line no-param-reassign
    jsonData = removeDuplicatedLinks(jsonData);

    let allErrors = '';
    let allWarnings = '';

    // Check if json data is a valid json array
    // Validate JSON
    let { errors, warnings } = validate(jsonData, 'import');

    warnings = [...warnings, ...warnIfUnrecognizedFieldsWillBeIgnored(jsonData)];

    allErrors = errors.join('<hr/>');
    allWarnings = warnings.join('<hr/>');

    if (allErrors.length > 0) {
      setDialogContent('Error!', allErrors);
      return;
    }

    if (allWarnings.length > 0) {
      setDialogContent('Warning!', allWarnings, () => {
        setDialogContent('Warning!', '<p>Do you want to import anyway?</p>', async () => {
          await importFileData(context, jsonData, setDialogContent, setIsImporting);
        });
      });
    } else {
      await importFileData(context, jsonData, setDialogContent, setIsImporting);
    }
  });
}

async function importFileData(
  context: Excel.RequestContext,
  jsonData: any,
  setDialogContent: any,
  setIsImporting: (v: boolean) => void
) {
  setDialogContent('Wait a moment...', 'Importing data...');
  setIsImporting(true);
  try {
    // Ignore types/classes that are not recognized
    const filteredItems = Array.isArray(jsonData)
      ? jsonData.filter((data: any) => Object.keys(map).includes(data['@type'].split(':')[1]))
      : [];
    await importByData(context, filteredItems);
  } catch (error: any) {
    setIsImporting(false);
    setDialogContent('Error!', error.message || 'Something went wrong');
    return;
  }
  setDialogContent('Success!', 'Your data has been successfully imported.');
  setIsImporting(false);
}

async function importByData(context: Excel.RequestContext, jsonData: any) {
  // Create Tables if they don't exist
  await createSheetsAndTables();

  // Write Simple Records to Tables
  await writeTable(context, jsonData);

  // Write Linked Records to Tables
  await writeTableLinked(context, jsonData);
}

async function writeTable(
  context: Excel.RequestContext,
  tableData: TableInterface[]
): Promise<void> {
  for (const data of tableData) {
    const tableName = data['@type'].split(':')[1];
    const recordId = data['@id'];
    const worksheet = context.workbook.worksheets.getItem(tableName);
    worksheet.load('tables');
    context.trackedObjects.add(worksheet);
    await context.sync();
    const table = worksheet.tables.getItem(tableName);
    const tableRange = table.getRange();
    const tableHeaderRange = table.getHeaderRowRange();
    tableRange.load('values');
    tableHeaderRange.load('values');
    context.trackedObjects.add(tableRange);
    context.trackedObjects.add(tableHeaderRange);
    await context.sync();
    const idColumnIndex = tableHeaderRange.values[0].indexOf('@id');
    const idColumn = tableRange.getColumn(idColumnIndex);
    idColumn.load('values');
    context.trackedObjects.add(idColumn);
    await context.sync();
    const idColumnValues = idColumn.values;

    // Create the record
    let record: { [key: string]: unknown } = {};
    Object.entries(data).forEach(([key, value]) => {
      let fieldName = key;
      if (
        !checkIfFieldIsRecognized(tableName, fieldName) &&
        fieldName !== '@type' &&
        fieldName !== '@context' &&
        fieldName !== 'value' &&
        fieldName !== 'hasLegalName'
      ) {
        return;
      }
      if (fieldName === 'value') {
        fieldName = 'i72:value';
      }
      if (fieldName === 'hasLegalName') {
        fieldName = 'org:hasLegalName';
      }
      const cid = new map[tableName as ModelType]();
      if (
        fieldName !== '@type' &&
        fieldName !== '@context' &&
        cid.getFieldByName(fieldName)?.type !== 'link' &&
        cid.getFieldByName(fieldName)?.type
      ) {
        if (cid.getFieldByName(fieldName)?.type === 'i72' || fieldName === 'value') {
          record[fieldName] =
            // @ts-ignore
            value?.numerical_value?.toString() || value?.['i72:numerical_value']?.toString();

          // Extract the unit_of_measure value
          const unit_of_measure =
            (value as any)?.['i72:unit_of_measure'] || (value as any)?.['unit_of_measure'];
          if (unit_of_measure) {
            record['i72:unit_of_measure'] = unit_of_measure;
          }
        } else {
          record[fieldName] = value;
        }
      }
    });

    // Add or Update the record on the table
    // check if the record already exists
    const idColumnValue = idColumnValues.map((item) => item[0].toString());
    const idIndex = idColumnValue.indexOf(recordId);
    let row: Excel.Range;
    if (idIndex !== -1) {
      row = tableRange.getRow(idIndex);
    } else {
      // Add the record
      // Get first row with empty id
      const emptyIdIndex = idColumnValue.indexOf('');
      row = tableRange.getRow(emptyIdIndex);
    }
    row.load('values');
    context.trackedObjects.add(row);
    await context.sync();

    // Update the record
    for (const [key, value] of Object.entries(record)) {
      const columnIndex = tableHeaderRange.values[0].indexOf(key);
      row.values = row.values || [];
      row.values[0][columnIndex] = value;
    }

    context.trackedObjects.remove(row);
    context.trackedObjects.remove(idColumn);
    context.trackedObjects.remove(tableRange);
    context.trackedObjects.remove(tableHeaderRange);
    context.trackedObjects.remove(worksheet);
    await context.sync();
  }
}

async function writeTableLinked(
  context: Excel.RequestContext,
  tableData: TableInterface[]
): Promise<void> {
  console.log('tableData', tableData);
  for (const data of tableData) {
    const tableName = data['@type'].split(':')[1];
    const recordId = data['@id'];
    const worksheet = context.workbook.worksheets.getItem(tableName);
    worksheet.load('tables');
    context.trackedObjects.add(worksheet);
    await context.sync();
    const table = worksheet.tables.getItem(tableName);
    const tableRange = table.getRange();
    const tableHeaderRange = table.getHeaderRowRange();
    tableRange.load('values');
    tableHeaderRange.load('values');
    context.trackedObjects.add(tableRange);
    context.trackedObjects.add(tableHeaderRange);
    await context.sync();
    const idColumnIndex = tableHeaderRange.values[0].indexOf('@id');
    const idColumn = tableRange.getColumn(idColumnIndex);
    idColumn.load('values');
    context.trackedObjects.add(idColumn);
    await context.sync();
    const idColumnValues = idColumn.values;

    // Create the record
    const record: { [key: string]: unknown } = {};
    for (let [key, value] of Object.entries(data)) {
      if (key === 'value') {
        key = 'i72:value';
      }
      if (key === 'hasLegalName') {
        key = 'org:hasLegalName';
      }
      const cid = new map[tableName as ModelType]();
      if (key !== '@type' && key !== '@context' && cid.getFieldByName(key)?.type === 'link') {
        // Skip if the value is empty
        if (!value) continue;

        // @ts-ignore
        if (!Array.isArray(value)) value = [value];

        // remove duplicates from value
        value = [...new Set(value as string[])];

        record[key] = value;
      }
    }

    // Add or Update the record on the table
    // check if the record already exists
    const idColumnValue = idColumnValues.map((item) => item[0].toString());
    let idIndex = idColumnValue.indexOf(recordId);
    let row: Excel.Range;
    if (idIndex !== -1) {
      row = tableRange.getRow(idIndex);
    } else {
      // Add the record
      // Get first row with empty id
      idIndex = idColumnValue.indexOf('');
      row = tableRange.getRow(idIndex);
    }
    row.load('values');
    context.trackedObjects.add(row);
    await context.sync();
    const rowValues = row.values[0];

    // Update the record
    for (let [key, value] of Object.entries(record)) {
      const columnIndex = tableHeaderRange.values[0].indexOf(key);
      value = [
        ...new Set([
          ...((rowValues[columnIndex] as string).split(', ') || []),
          ...((value as string[]) || []),
        ]),
      ].filter((v) => v !== null && v !== undefined && v !== '');
      row.getCell(0, columnIndex).values = [[(value as string[]).join(', ')]];
    }

    context.trackedObjects.remove(row);
    await context.sync();

    // Update the linked tables
    for (const [key, value] of Object.entries(record)) {
      const relatedTableName = key.substring(3);
      const relatedFieldName = key.startsWith('has') ? `for${tableName}` : `has${tableName}`;
      await updateLinkedTablesFields(
        context,
        recordId,
        value as string[],
        relatedTableName,
        relatedFieldName
      );
    }

    context.trackedObjects.remove(idColumn);
    context.trackedObjects.remove(tableRange);
    context.trackedObjects.remove(tableHeaderRange);
    context.trackedObjects.remove(worksheet);
    await context.sync();
  }
}

async function updateLinkedTablesFields(
  context: Excel.RequestContext,
  currentFieldId: string,
  currentFiledValues: string[],
  relatedTableName: string,
  relatedFieldName: string
) {
  const worksheet = context.workbook.worksheets.getItem(relatedTableName);
  worksheet.load('tables');
  context.trackedObjects.add(worksheet);
  await context.sync();
  const table = worksheet.tables.getItem(relatedTableName);
  const tableHeadersRange = table.getHeaderRowRange();
  const tableRange = table.getRange();
  tableHeadersRange.load('values');
  tableRange.load('values');
  context.trackedObjects.add(tableHeadersRange);
  context.trackedObjects.add(tableRange);
  await context.sync();
  const relatedFieldIndex = tableHeadersRange.values[0].indexOf(relatedFieldName);
  const idColumnIndex = tableHeadersRange.values[0].indexOf('@id');
  const relatedFieldColumn = tableRange.getColumn(relatedFieldIndex);
  const idColumn = tableRange.getColumn(idColumnIndex);
  relatedFieldColumn.load('values');
  idColumn.load('values');
  context.trackedObjects.add(relatedFieldColumn);
  context.trackedObjects.add(idColumn);
  await context.sync();
  const relatedFieldValues = relatedFieldColumn.values;
  const idColumnValues = idColumn.values;

  // For each cell in the related field column, check if the id is in the array
  for (let i = 0; i < relatedFieldValues.length; i++) {
    const idColumnValue = idColumnValues[i][0].toString();
    const relatedFieldValue = relatedFieldValues[i][0].toString();
    const relatedFieldValueArray: string[] = relatedFieldValue.split(', ');

    if (!idColumnValue) {
      continue;
    }

    if (
      !currentFiledValues.includes(idColumnValue) &&
      relatedFieldValueArray.includes(currentFieldId)
    ) {
      // Remove the id from the array
      const newValueArray = relatedFieldValueArray.filter((v: string) => v !== currentFieldId);
      const newValue = newValueArray.join(', ');
      relatedFieldColumn.getCell(i, 0).values = [[newValue]];
    } else if (
      currentFiledValues.includes(idColumnValue) &&
      !relatedFieldValueArray.includes(currentFieldId)
    ) {
      // Add the id to the array
      relatedFieldColumn.getCell(i, 0).values = [
        [
          relatedFieldValueArray
            .concat(currentFieldId)
            .filter((v) => v !== null && v !== undefined && v !== '')
            .join(', '),
        ],
      ];
    }
  }

  context.trackedObjects.remove(relatedFieldColumn);
  context.trackedObjects.remove(idColumn);
  context.trackedObjects.remove(tableRange);
  context.trackedObjects.remove(tableHeadersRange);
  context.trackedObjects.remove(worksheet);
  await context.sync();
}

function removeDuplicatedLinks(jsonData: any) {
  for (const data of jsonData) {
    for (const [key, value] of Object.entries(data)) {
      if (Array.isArray(value)) {
        data[key] = [...new Set(value)];
      }
    }
  }
  return jsonData;
}

function validateIfEmptyFile(tableData: TableInterface[]) {
  if (!Array.isArray(tableData) || tableData.length === 0) {
    return true;
  }
  return false;
}

function doAllRecordsHaveId(tableData: TableInterface[]) {
  for (const data of tableData) {
    if (data['@id'] === undefined) {
      return false;
    }
  }
  return true;
}

function warnIfUnrecognizedFieldsWillBeIgnored(tableData: TableInterface[]) {
  const warnings = [];
  const classesSet = new Set();
  for (const data of tableData) {
    const tableName = data['@type'].split(':')[1];
    if (!Object.keys(map).includes(tableName)) {
      continue;
    }
    if (classesSet.has(tableName)) {
      continue;
    }
    classesSet.add(tableName);
    for (const key in data) {
      if (
        !checkIfFieldIsRecognized(tableName, key) &&
        key !== '@type' &&
        key !== '@context' &&
        key !== 'value' &&
        key !== 'hasLegalName'
      ) {
        warnings.push(
          `Table <b>${tableName}</b> has unrecognized field <b>${key}</b>. This field will be ignored.`
        );
      }
    }
  }
  return warnings;
}

function checkIfFieldIsRecognized(tableName: string, fieldName: string) {
  if (tableName === Organization.className && fieldName === 'hasLegalName') return true;
  if (tableName === IndicatorReport.className && fieldName === 'value') return true;
  const cid = new map[tableName as ModelType]();
  return cid
    .getFields()
    .map((item) => item.name)
    .includes(fieldName);
}
