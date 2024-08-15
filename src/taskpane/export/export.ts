import { TableInterface } from '../domain/interfaces/table.interface';
import { ignoredFields, map, ModelType } from '../domain/models';
import { Base as BaseModel } from '../domain/models/Base';
import { validate } from '../domain/validation/validator';
import { downloadJSONLD } from '../utils/utils';

/* global Excel */
export async function exportData(
  orgName: string,
  setDialogContent: (header: string, content: string, nextCallBack?: Function) => void
): Promise<void> {
  await Excel.run(async (context: Excel.RequestContext) => {
    // Get the tables from the workbook
    const workbook = context.workbook;
    workbook.load('tables');
    await context.sync();
    const tables = workbook.tables.items;

    const data: TableInterface[] = [];

    const tableNames = tables.map((item) => item.name);
    for (const [key] of Object.entries(map)) {
      if (!tableNames.includes(key)) {
        setDialogContent(
          'Error!',
          `Table <b>${key}</b> is missing. Please create the tables first.`
        );
        return;
      }
    }

    for (const table of tables) {
      // If the table is not in the map, skip it
      if (!Object.keys(map).includes(table.name)) {
        continue;
      }

      // Get the records from the table
      const tableRange = table.getRange();
      const tableHeaderRange = tableRange.getRow(0);
      tableHeaderRange.load('values');
      table.load('values, rows');
      await context.sync();
      const records = table.rows.items;

      const cid: BaseModel = new map[table.name as ModelType]();
      for (const record of records) {
        record.load('values');
        await context.sync();

        const row: TableInterface = {
          '@context': 'http://ontology.commonapproach.org/contexts/cidsContext.json',
          '@type': `cids:${table.name}`,
          '@id': '',
        };

        let isEmpty = true; // Flag to check if the row is empty
        for (const field of cid.getFields()) {
          const columnIndex = tableHeaderRange.values[0].indexOf(field.name);
          const value: any = record.values[0][columnIndex];
          if (field.type === 'link') {
            if (field.representedType === 'array') {
              const fieldValue = value ?? field?.defaultValue;
              if (fieldValue && fieldValue.length > 0) {
                isEmpty = false;
              }
              row[field.name] =
                typeof fieldValue === 'string'
                  ? fieldValue.split(', ').filter((v) => v !== '' && v !== null && v !== undefined)
                  : (fieldValue as string[]).filter(
                      (v) => v !== '' && v !== null && v !== undefined
                    );
            } else if (field.representedType === 'string') {
              const fieldValue = value ?? field?.defaultValue;
              if (fieldValue) {
                isEmpty = false;
              }
              row[field.name] = Array.isArray(fieldValue) ? fieldValue[0] : fieldValue;
            }
          } else if (field.type === 'i72') {
            if (field.name === 'i72:value') {
              const numericalValue = value ?? field?.defaultValue;
              const unitOfMeasureColumnIndex =
                tableHeaderRange.values[0].indexOf('i72:unit_of_measure');
              const unitOfMeasure = record.values[0][unitOfMeasureColumnIndex] ?? '';
              if (numericalValue || unitOfMeasure) {
                isEmpty = false;
              }
              row[field.name] = {
                '@context': 'http://ontology.commonapproach.org/contexts/cidsContext.json',
                '@type': 'i72:Measure',
                'i72:numerical_value': numericalValue.toString(),
                'i72:unit_of_measure': unitOfMeasure.toString(),
              };
            }
          } else {
            const fieldValue = value ?? '';
            if (fieldValue) {
              isEmpty = false;
            }
            row[field.name] = fieldValue.toString();
          }
        }
        if (!isEmpty) {
          data.push(row);
        }
      }
    }

    const { errors, warnings } = validate(data, 'export');

    const noExportingFields = await checkForNotExportedFields(context);
    const emptyTableWarning = await checkForEmptyTables(context);
    const allWarnings = noExportingFields + warnings.join('<hr/>') + emptyTableWarning;

    if (errors.length > 0) {
      setDialogContent('Error!', errors.map((item) => `<p>${item}</p>`).join(''));
      return;
    }

    if (allWarnings.length > 0) {
      setDialogContent('Warning!', allWarnings, () => {
        setDialogContent('Warning!', '<p>Do you want to export anyway?</p>', () => {
          downloadJSONLD(data, `${getFileName(orgName)}.json`);
          setDialogContent('Success!', 'Data exported successfully!');
        });
      });
      return;
    }
    downloadJSONLD(data, `${getFileName(orgName)}.json`);
    setDialogContent('Success!', 'Data exported successfully!');
  });
}

function getFileName(orgName: string): string {
  const date = new Date();

  // Get the year, month, and day from the date
  const year = date.getFullYear();
  const month = date.getMonth() + 1; // Add 1 because months are 0-indexed.
  const day = date.getDate();

  // Format month and day to ensure they are two digits
  const monthFormatted = month < 10 ? '0' + month : month;
  const dayFormatted = day < 10 ? '0' + day : day;

  // Concatenate the components to form the desired format (YYYYMMDD)
  const timestamp = `${year}${monthFormatted}${dayFormatted}`;

  return `CIDSBasic${orgName}${timestamp}`;
}

async function checkForNotExportedFields(context: Excel.RequestContext) {
  const workbook = context.workbook;
  workbook.load('tables');
  await context.sync();
  const tables = workbook.tables.items;

  let warnings = '';
  for (const table of tables) {
    if (!Object.keys(map).includes(table.name)) {
      continue;
    }
    const cid = new map[table.name as ModelType]();
    const internalFields = cid.getFields().map((item) => item.name);

    const tableHeaderRange = table.getHeaderRowRange();
    tableHeaderRange.load('values');
    await context.sync();
    const tableHeadersValues = tableHeaderRange.values[0];

    for (const field of tableHeadersValues) {
      if (Object.keys(map).includes(field) || ignoredFields[table.name]?.includes(field)) {
        continue;
      }
      if (!internalFields.includes(field)) {
        warnings += `Field <b>${field}</b> on table <b>${table.name}</b> will not be exported<hr/>`;
      }
    }
  }
  return warnings;
}

async function checkForEmptyTables(context: Excel.RequestContext) {
  const workbook = context.workbook;
  workbook.load('tables');
  await context.sync();
  const tables = workbook.tables.items;

  let warnings = '';
  for (const table of tables) {
    if (!Object.keys(map).includes(table.name)) {
      continue;
    }

    const tableDataRange = table.getDataBodyRange();
    tableDataRange.load('values');
    await context.sync();
    const tableData = tableDataRange.values;

    let isEmpty = true;
    for (const row of tableData) {
      for (const cell of row) {
        if (cell) isEmpty = false;
      }
    }

    if (isEmpty) {
      warnings += `<hr/>Table <b>${table.name}</b> is empty<hr/>`;
    }
  }
  return warnings;
}
