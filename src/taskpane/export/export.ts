// import Base from '@airtable/blocks/dist/types/src/models/base';
// import { IntlShape } from 'react-intl';
// import { LinkedCellInterface } from '../domain/interfaces/cell.interface';
// import { ignoredFields, map } from '../domain/models';
// import { Base as BaseModel } from '../domain/models/Base';
// import { validate } from '../domain/validation/validator';
// import { downloadJSONLD } from '../utils';
// export async function exportData(
//   base: Base,
//   setDialogContent: (
//     header: string,
//     text: string,
//     open: boolean,
//     nextCallback?: () => void
//   ) => void,
//   orgName: string,
//   intl: IntlShape
// ): Promise<void> {
//   const tables = base.tables;
//   let data = [];

//   const tableNames = tables.map((item) => item.name);
//   for (const [key] of Object.entries(map)) {
//     if (!tableNames.includes(key)) {
//       setDialogContent(
//         `${intl.formatMessage({
//           id: 'generics.error',
//           defaultMessage: 'Error',
//         })}!`,
//         intl.formatMessage(
//           {
//             id: 'export.messages.error.missingTable',
//             defaultMessage: `Table <b>{tableName}</b> is missing. Please create the tables first.`,
//           },
//           { tableName: key, b: (str) => `<b>${str}</b>` }
//         ),
//         true
//       );
//       return;
//     }
//   }

//   for (const table of tables) {
//     // If the table is not in the map, skip it
//     if (!Object.keys(map).includes(table.name)) {
//       continue;
//     }

//     const records = (await table.selectRecordsAsync()).records;

//     const cid: BaseModel = new map[table.name]();
//     for (const record of records) {
//       let row = {
//         '@context': 'http://ontology.commonapproach.org/contexts/cidsContext.json',
//         '@type': `cids:${table.name}`,
//       };
//       let isEmpty = true; // Flag to check if the row is empty
//       for (const field of cid.getFields()) {
//         if (field.type === 'link') {
//           const value: any = record.getCellValue(field.name);
//           if (field.representedType === 'array') {
//             const fieldValue =
//               value?.map((item: LinkedCellInterface) => item.name) ?? field?.defaultValue;
//             if (fieldValue && fieldValue.length > 0) {
//               isEmpty = false;
//             }
//             row[field.name] = fieldValue;
//           } else if (field.representedType === 'string') {
//             const fieldValue = value ? value[0]?.name : field?.defaultValue;
//             if (fieldValue) {
//               isEmpty = false;
//             }
//             row[field.name] = fieldValue;
//           }
//         } else if (field.type === 'i72') {
//           if (field.name === 'i72:value') {
//             const numericalValue = record.getCellValueAsString(field.name) ?? field?.defaultValue;
//             const unitOfMeasure = record.getCellValueAsString('i72:unit_of_measure') ?? '';
//             if (numericalValue || unitOfMeasure) {
//               isEmpty = false;
//             }
//             row[field.name] = {
//               '@context': 'http://ontology.commonapproach.org/contexts/cidsContext.json',
//               '@type': 'i72:Measure',
//               'i72:numerical_value': numericalValue,
//               'i72:unit_of_measure': unitOfMeasure,
//             };
//           }
//         } else {
//           const fieldValue = record.getCellValue(field.name) ?? '';
//           if (fieldValue) {
//             isEmpty = false;
//           }
//           row[field.name] = fieldValue;
//         }
//       }
//       if (!isEmpty) {
//         data.push(row);
//       }
//     }
//   }

//   const { errors, warnings } = validate(data, 'export', intl);

//   const emptyTableWarning = await checkForEmptyTables(base, intl);
//   const allWarnings =
//     checkForNotExportedFields(base, intl) + warnings.join('<hr/>') + emptyTableWarning;

//   if (errors.length > 0) {
//     setDialogContent(
//       `${intl.formatMessage({
//         id: 'generics.error',
//         defaultMessage: 'Error',
//       })}!`,
//       errors.map((item) => `<p>${item}</p>`).join(''),
//       true
//     );
//     return;
//   }

//   if (allWarnings.length > 0) {
//     setDialogContent(
//       `${intl.formatMessage({
//         id: 'generics.warning',
//         defaultMessage: 'Warning',
//       })}!`,
//       allWarnings,
//       true,
//       () => {
//         setDialogContent(
//           `${intl.formatMessage({
//             id: 'generics.warning',
//             defaultMessage: 'Warning',
//           })}!`,
//           intl.formatMessage({
//             id: 'export.messages.warning.continue',
//             defaultMessage: '<p>Do you want to export anyway?</p>',
//           }),
//           true,
//           () => {
//             downloadJSONLD(data, `${getFileName(orgName)}.json`);
//             setDialogContent('', '', false);
//           }
//         );
//       }
//     );
//     return;
//   }
//   downloadJSONLD(data, `${getFileName(orgName)}.json`);
// }

// function getFileName(orgName: string): string {
//   const date = new Date();

//   // Get the year, month, and day from the date
//   const year = date.getFullYear();
//   const month = date.getMonth() + 1; // Add 1 because months are 0-indexed.
//   const day = date.getDate();

//   // Format month and day to ensure they are two digits
//   const monthFormatted = month < 10 ? '0' + month : month;
//   const dayFormatted = day < 10 ? '0' + day : day;

//   // Concatenate the components to form the desired format (YYYYMMDD)
//   const timestamp = `${year}${monthFormatted}${dayFormatted}`;

//   return `CIDSBasic${orgName}${timestamp}`;
// }

// function checkForNotExportedFields(base: Base, intl: IntlShape) {
//   let warnings = '';
//   for (const table of base.tables) {
//     if (!Object.keys(map).includes(table.name)) {
//       continue;
//     }
//     const cid = new map[table.name]();
//     const internalFields = cid.getFields().map((item) => item.name);
//     const externalFields = table.fields.map((item) => item.name);

//     for (const field of externalFields) {
//       if (Object.keys(map).includes(field) || ignoredFields[table.name]?.includes(field)) {
//         continue;
//       }
//       if (!internalFields.includes(field)) {
//         warnings += intl.formatMessage(
//           {
//             id: 'export.messages.warning.fieldWillNotBeExported',
//             defaultMessage: `Field <b>{fieldName}</b> on table <b>{tableName}</b> will not be exported<hr/>`,
//           },
//           {
//             fieldName: field,
//             tableName: table.name,
//             b: (str) => `<b>${str}</b>`,
//           }
//         );
//       }
//     }
//   }
//   return warnings;
// }

// async function checkForEmptyTables(base: Base, intl: IntlShape) {
//   let warnings = '';
//   for (const table of base.tables) {
//     if (!Object.keys(map).includes(table.name)) {
//       continue;
//     }
//     const records = await table.selectRecordsAsync();
//     if (records.records.length === 0) {
//       warnings += intl.formatMessage(
//         {
//           id: 'export.messages.warning.emptyTable',
//           defaultMessage: `<hr/>Table <b>{tableName}</b> is empty<hr/>`,
//         },
//         {
//           tableName: table.name,
//           b: (str) => `<b>${str}</b>`,
//         }
//       );
//     }
//   }
//   return warnings;
// }
