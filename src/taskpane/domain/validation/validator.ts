import moment from "moment";
import { IntlShape } from "react-intl";
import { getCodeListByTableName } from "../fetchServer/getCodeLists";
import { getContext } from "../fetchServer/getContext";
import { TableInterface } from "../interfaces/table.interface";
import {
  contextUrl,
  map,
  mapSFFModel,
  ModelType,
  predefinedCodeLists,
  SFFModelType,
} from "../models";
import { FieldType } from "../models/Base";

type Operation = "import" | "export";

const validatorErrors = new Set<string>();
const validatorWarnings = new Set<string>();

export async function validate(
  tableData: TableInterface[],
  operation: Operation = "export",
  intl: IntlShape
): Promise<{
  errors: string[];
  warnings: string[];
}> {
  validatorWarnings.clear();
  validatorErrors.clear();

  validateIfEmptyFile(tableData, intl);

  validateIfIdIsValidUrl(tableData, operation, intl);

  // eslint-disable-next-line no-param-reassign
  tableData = removeEmptyRows(tableData);

  tableData.forEach((item) => {
    validateTypeProp(item, intl);
  });

  await validateRecords(tableData, operation, intl);

  return {
    errors: Array.from(validatorErrors),
    warnings: Array.from(validatorWarnings),
  };
}

async function validateRecords(tableData: TableInterface[], operation: Operation, intl: IntlShape) {
  // Records to keep track of unique values
  const uniqueRecords: Record<string, Set<any>> = {};

  await validateLinkedFields(tableData, operation, intl);

  for (const data of tableData) {
    if (validateTypeProp(data, intl)) return;
    const tableName = data["@type"].split(":")[1];
    const id = data["@id"];

    // Initialize the schema for the table
    let cid;
    // Check if type is part of the SFF module
    if (mapSFFModel[tableName as SFFModelType]) {
      cid = new mapSFFModel[tableName as SFFModelType]();
    } else {
      cid = new map[tableName as ModelType]();
    }

    // Initialize a record for this table if not already present
    if (!uniqueRecords[tableName]) {
      uniqueRecords[tableName] = new Set();
    }

    //check if required fields are present
    for (const field of cid.getAllFields()) {
      if (
        field.required &&
        !Object.keys(data)
          .map((d) => (d.indexOf(":") !== -1 ? d.split(":")[1] : d))
          .includes(field.name.indexOf(":") !== -1 ? field.name.split(":")[1] : field.name)
      ) {
        if (operation === "import" && field.name !== "@id") {
          validatorWarnings.add(
            intl.formatMessage(
              {
                id: "validation.messages.missingRequiredField",
                defaultMessage:
                  "Required field <b>{fieldName}</b> is missing on table <b>{tableName}</b>",
              },
              {
                fieldName: field.displayName || field.name,
                tableName,
                b: (str) => `<b>${str}</b>`,
              }
            )
          );
        } else {
          validatorErrors.add(
            intl.formatMessage(
              {
                id: "validation.messages.missingRequiredField",
                defaultMessage:
                  "Required field <b>{fieldName}</b> is missing on table <b>{tableName}</b>",
              },
              {
                fieldName: field.displayName || field.name,
                tableName,
                b: (str) => `<b>${str}</b>`,
              }
            )
          );
        }
      }
    }

    for (const field of cid.getAllFields()) {
      if (field.semiRequired) {
        if (
          !Object.keys(data)
            .map((d) => (d.indexOf(":") !== -1 ? d.split(":")[1] : d))
            .includes(field.name.indexOf(":") !== -1 ? field.name.split(":")[1] : field.name)
        ) {
          validatorWarnings.add(
            intl.formatMessage(
              {
                id: "validation.messages.missingRequiredField",
                defaultMessage:
                  "Required field <b>{fieldName}</b> is missing on table <b>{tableName}</b>",
              },
              {
                fieldName: field.displayName || field.name,
                tableName,
                b: (str) => `<b>${str}</b>`,
              }
            )
          );
        }
        // @ts-ignore
        if (data[field.name]?.length === 0) {
          validatorWarnings.add(
            intl.formatMessage(
              {
                id: "validation.messages.emptyField",
                defaultMessage: "Field <b>{fieldName}</b> is empty on table <b>{tableName}</b>",
              },
              {
                fieldName: field.displayName || field.name,
                tableName,
                b: (str) => `<b>${str}</b>`,
              }
            )
          );
        }
      }
    }

    // check if notNull fields are not null
    for (const field of cid.getAllFields()) {
      if (
        field.notNull &&
        operation === "export" &&
        ((!data[field.name] && !data[field.name.split(":")[1]]) ||
          isFieldValueNullOrEmpty(data[field.name] || data[field.name.split(":")[1]]))
      ) {
        const msg = intl
          .formatMessage(
            {
              id: "validation.messages.nullOrEmptyField",
              defaultMessage:
                "Field <b>{fieldName}</b> is null or empty on table <b>{tableName}</b>",
            },
            {
              fieldName: field.displayName || field.name,
              tableName,
              b: (str) => `<b>${str}</b>`,
            }
          )
          .toString();
        validatorErrors.add(msg);
      }
    }

    for (let [fieldName, fieldValue] of Object.entries(data)) {
      if (fieldName === "@context" || fieldName === "@type") continue;
      let fieldProps: FieldType | null = null;
      try {
        fieldProps = cid.getFieldByName(fieldName);
      } catch (_) {
        continue;
      }

      if (!fieldProps) {
        continue;
      }

      const tableFields = cid.getAllFields().map((field) => field.name);
      const fieldDisplayName = cid.getFieldByName(fieldName)?.displayName || fieldName;

      for (const field of tableFields) {
        if (field.indexOf(":") !== -1) {
          const splitField = field.split(":");
          if (splitField[1] === fieldName) {
            fieldName = field;
            break;
          }
        }
      }

      if (Array.isArray(fieldValue)) {
        // check if fieldValue has duplicate values
        const uniqueValues = new Set(fieldValue);
        if (uniqueValues.size !== fieldValue.length) {
          validatorWarnings.add(
            intl.formatMessage(
              {
                id: "validation.messages.duplicateFieldValues",
                defaultMessage:
                  "Duplicate values in field <b>{fieldName}</b> on table <b>{tableName}</b>",
              },
              {
                fieldName: fieldDisplayName,
                tableName,
                b: (str) => `<b>${str}</b>`,
              }
            )
          );
        }
      }

      if (fieldProps.type !== "object") {
        // Validate unique fields
        if (fieldProps?.unique) {
          const uniqueResult = validateUnique(
            tableName,
            fieldName,
            fieldValue,
            uniqueRecords,
            id,
            intl
          );
          if (!uniqueResult.isUnique && uniqueResult.reason === "duplicate") {
            const msg = intl
              .formatMessage(
                {
                  id: "validation.messages.duplicateUniqueFieldValue",
                  defaultMessage:
                    "Duplicate value for unique field <b>{fieldName}</b>: <b>{fieldValue}</b> in table <b>{tableName}</b>",
                },
                {
                  fieldName: fieldDisplayName,
                  fieldValue: fieldValue ? fieldValue.toString() : "null",
                  tableName,
                  b: (str) => `<b>${str}</b>`,
                }
              )
              .toString();
            if (fieldName !== "@id") {
              validatorWarnings.add(msg);
            } else {
              validatorErrors.add(msg);
            }
          }
        }

        if (fieldProps?.notNull) {
          if (fieldValue === "" || !fieldValue) {
            validatorWarnings.add(
              intl.formatMessage(
                {
                  id: "validation.messages.warning.nullOrEmptyField",
                  defaultMessage:
                    "Field <b>{fieldName}</b> on table <b>{tableName}</b> is null or empty.",
                },
                {
                  fieldName: fieldDisplayName,
                  tableName,
                  b: (str) => `<b>${str}</b>`,
                }
              )
            );
          }
        }

        if (fieldProps?.required) {
          if (fieldValue === "" || !fieldValue) {
            validatorWarnings.add(
              intl.formatMessage(
                {
                  id: "validation.messages.warning.missingRequiredField",
                  defaultMessage:
                    "Field <b>{fieldName}</b> on table <b>{tableName}</b> is required.",
                },
                {
                  fieldName: fieldDisplayName,
                  tableName,
                  b: (str) => `<b>${str}</b>`,
                }
              )
            );
          }
        }

        if (fieldProps?.type === "select") {
          if (fieldProps.selectOptions || fieldProps.getOptionsAsync) {
            let shouldWarn = false;
            if (fieldProps.getOptionsAsync) {
              const options = await fieldProps.getOptionsAsync();
              if (
                !options.find(
                  (op) => op.id === (Array.isArray(fieldValue) ? fieldValue[0] : fieldValue)
                )
              ) {
                shouldWarn = true;
              }
            } else if (
              !fieldProps.selectOptions?.find(
                (op) => op.id === (Array.isArray(fieldValue) ? fieldValue[0] : fieldValue)
              )
            ) {
              shouldWarn = true;
            }
            if (shouldWarn) {
              validatorWarnings.add(
                intl.formatMessage(
                  {
                    id: "validation.messages.warning.invalidSelectField",
                    defaultMessage:
                      "Field <b>{fieldName}</b> on table <b>{tableName}</b> has an invalid value.",
                  },
                  {
                    fieldName: fieldDisplayName,
                    tableName,
                    b: (str) => `<b>${str}</b>`,
                  }
                )
              );
            }
          }
        }

        if (fieldProps?.type === "multiselect") {
          if (fieldProps.selectOptions || fieldProps.getOptionsAsync) {
            let shouldWarn = false;
            const selectedValues = Array.isArray(fieldValue) ? fieldValue : [fieldValue];
            if (fieldProps.getOptionsAsync) {
              const options = await fieldProps.getOptionsAsync();
              selectedValues.forEach((value) => {
                if (!options.find((op) => op.id === value)) {
                  shouldWarn = true;
                }
              });
            } else {
              selectedValues.forEach((value) => {
                if (!fieldProps.selectOptions?.find((op) => op.id === value)) {
                  shouldWarn = true;
                }
              });
            }
            if (shouldWarn) {
              validatorWarnings.add(
                intl.formatMessage(
                  {
                    id: "validation.messages.warning.invalidSelectField",
                    defaultMessage:
                      "Field <b>{fieldName}</b> on table <b>{tableName}</b> has invalid values.",
                  },
                  {
                    fieldName: fieldDisplayName,
                    tableName,
                    b: (str) => `<b>${str}</b>`,
                  }
                )
              );
            }
          }
        }

        // Validate field values against context for basic types
        // xsd:string, xsd:anyURI, xsd:nonNegativeNumber, xsd:boolean, xsd:date
        // If the the field is in the context and has a @type property of one of this types
        // we need to validate the value against the type, e.g if the field is a xsd:boolean the value should be true or false
        const context: any = mapSFFModel[tableName as SFFModelType]
          ? await getContext(contextUrl[1])
          : await getContext(contextUrl[0]);
        const fieldContext = context["@context"][fieldName];
        if (fieldContext) {
          const fieldType = fieldContext["@type"];
          if (fieldType) {
            const value = fieldValue;
            if (value && fieldType === "xsd:string" && typeof value !== "string") {
              validatorWarnings.add(
                intl.formatMessage(
                  {
                    id: "validation.messages.warning.invalidStringType",
                    defaultMessage:
                      "The field <b>{fieldName}</b> in <b>{tableName}</b> must be text.",
                  },
                  {
                    fieldName: fieldDisplayName,
                    tableName,
                    b: (str) => `<b>${str}</b>`,
                  }
                )
              );
              // set field value to default value
              data[fieldName] = "";
            } else if (value && fieldType === "xsd:anyURI") {
              try {
                if (typeof value !== "string") throw new Error();
                new URL(value);
              } catch (error) {
                validatorWarnings.add(
                  intl.formatMessage(
                    {
                      id: "validation.messages.warning.invalidUrlType",
                      defaultMessage:
                        "The field <b>{fieldName}</b> in <b>{tableName}</b> must be a valid URL.",
                    },
                    {
                      fieldName: fieldDisplayName,
                      tableName,
                      b: (str) => `<b>${str}</b>`,
                    }
                  )
                );
              }
            } else if (
              value &&
              (fieldType === "xsd:nonNegativeNumber" || fieldType === "xsd:nonNegativeInteger")
            ) {
              const numValue = +value;
              if (isNaN(numValue) || !Number.isInteger(numValue) || numValue < 0) {
                validatorWarnings.add(
                  intl.formatMessage(
                    {
                      id: "validation.messages.warning.invalidNumberType",
                      defaultMessage:
                        "The field <b>{fieldName}</b> in <b>{tableName}</b> must be a whole number greater than or equal to 0.",
                    },
                    {
                      fieldName: fieldDisplayName,
                      tableName,
                      b: (str) => `<b>${str}</b>`,
                    }
                  )
                );
              }
            } else if (fieldType === "xsd:boolean" && typeof value !== "boolean") {
              validatorWarnings.add(
                intl.formatMessage(
                  {
                    id: "validation.messages.warning.invalidBooleanType",
                    defaultMessage:
                      "The field <b>{fieldName}</b> in <b>{tableName}</b> must be either true or false.",
                  },
                  {
                    fieldName: fieldDisplayName,
                    tableName,
                    b: (str) => `<b>${str}</b>`,
                  }
                )
              );
              // set field value to default value
              data[fieldName] = false;
            } else if (value && fieldType === "xsd:date") {
              const validDateFormats = ["YYYY-MM-DD", "YYYY-MM-DDTHH:mm:ssZ"];
              const isValidDate = moment(value.toString(), validDateFormats, true).isValid();

              if (!isValidDate) {
                validatorWarnings.add(
                  intl.formatMessage(
                    {
                      id: "validation.messages.warning.invalidDateFormat",
                      defaultMessage:
                        "The field <b>{fieldName}</b> in <b>{tableName}</b> must be a valid date in format YYYY-MM-DD or YYYY-MM-DDTHH:mm:ssZ.",
                    },
                    {
                      fieldName: fieldDisplayName,
                      tableName,
                      b: (str) => `<b>${str}</b>`,
                    }
                  )
                );
                // set field value to default value
                data[fieldName] = "";
              }
            }
          }
        }
      }
    }
  }
}

function validateUnique(
  tableName: string,
  fieldName: string,
  fieldValue: any,
  uniqueRecords: Record<string, Set<any>>,
  id: string,
  intl: IntlShape
): { isUnique: boolean; reason?: "invalidUrl" | "duplicate" } {
  // Unique key for this field in the format "tableName.fieldName"
  if (!id) return { isUnique: false, reason: "duplicate" };
  let urlObject;

  try {
    urlObject = new URL(id);
  } catch (error) {
    validatorErrors.add(
      intl.formatMessage(
        {
          id: "validation.messages.invalidIdFormat",
          defaultMessage:
            "Invalid URL format: <b>{id}</b> for <b>@id</b> on table <b>{tableName}</b>",
        },
        {
          id,
          tableName,
          b: (str) => `<b>${str}</b>`,
        }
      )
    );
    return { isUnique: false, reason: "invalidUrl" };
  }

  const baseUrl = `${urlObject.protocol}//${urlObject.hostname}`;

  const uniqueKey = `${tableName}.${fieldName}.${baseUrl}`;

  // Initialize a record for this field if not already present
  if (!uniqueRecords[uniqueKey]) {
    // eslint-disable-next-line no-param-reassign
    uniqueRecords[uniqueKey] = new Set();
  }

  // Check if the value already exists
  if (uniqueRecords[uniqueKey].has(fieldValue)) {
    // Value is not unique
    return { isUnique: false, reason: "duplicate" };
  } else {
    // Record this value as encountered and return true
    uniqueRecords[uniqueKey].add(fieldValue);
    return { isUnique: true };
  }
}

function validateTypeProp(data: any, intl: IntlShape): boolean {
  if (!("@type" in data)) {
    validatorErrors.add(
      intl.formatMessage({
        id: "validation.messages.missingTypeProperty",
        defaultMessage: "<b>@type</b> must be present in the data",
      })
    );
    return true;
  }
  if (data["@type"].length === 0) {
    validatorErrors.add(
      intl.formatMessage({
        id: "validation.messages.emptyTypeProperty",
        defaultMessage: "<b>@type</b> cannot be empty",
      })
    );
    return true;
  }
  try {
    if (data["@type"]?.split(":")[1].length === 0) {
      validatorErrors.add(
        intl.formatMessage({
          id: "validation.messages.invalidTypeProperty",
          defaultMessage: "<b>@type</b> must follow the format <b>cids:tableName</b>",
        })
      );
      return true;
    }
  } catch (error) {
    validatorErrors.add(
      intl.formatMessage({
        id: "validation.messages.invalidTypeProperty",
        defaultMessage: "<b>@type</b> must follow the format <b>cids:tableName</b>",
      })
    );
    return true;
  }
  const tableName = (data["@type"] as string)?.split(":")[1];
  if (!map[tableName as ModelType] && !mapSFFModel[tableName as SFFModelType]) {
    validatorWarnings.add(
      intl.formatMessage(
        {
          id: "validation.messages.unrecognizedTypeProperty",
          defaultMessage:
            "Table <b>{tableName}</b> is not recognized in the basic tier and will be ignored.",
        },
        {
          tableName,
          b: (str) => `<b>${str}</b>`,
        }
      )
    );
    return true;
  }
  return false;
}

async function validateLinkedFields(
  tableData: TableInterface[],
  operation: Operation,
  intl: IntlShape
) {
  for (const data of tableData) {
    if (validateTypeProp(data, intl)) return;
    const tableName = data["@type"].split(":")[1];

    // Initialize the schema for the table
    let cid;
    // Check if type is part of the SFF module
    if (mapSFFModel[tableName as SFFModelType]) {
      cid = new mapSFFModel[tableName as SFFModelType]();
    } else {
      cid = new map[tableName as ModelType]();
    }

    // for each field that has type link, check if all linked ids exists
    const fields = cid.getAllFields();
    const linkedFields = fields.filter((field) => field.type === "link");
    linkedFields.forEach(async (field) => {
      const fieldName = field.name;
      if (!data[fieldName]) {
        data[fieldName] = [];
      }

      let isString = false;
      if (!Array.isArray(data[fieldName])) {
        if (typeof data[fieldName] === "string") {
          isString = true;
        }
        data[fieldName] =
          typeof data[fieldName] === "string" && data[fieldName].length > 0
            ? [...data[fieldName].split(", ")]
            : [];
      }

      if (data[fieldName].length === 0) {
        const msg = intl
          .formatMessage(
            {
              id: "validation.messages.missingLinkedField",
              defaultMessage: "{tableName} <b>{name}</b> has no {fieldName}",
            },
            {
              tableName,
              name: (data["org:hasLegalName"] || data["hasLegalName"] || data["hasName"]) as string,
              fieldName,
              b: (str) => `<b>${str}</b>`,
            }
          )
          .toString();
        if (field.required && operation === "export") {
          validatorErrors.add(msg);
        } else if (field.required || field.semiRequired) {
          validatorWarnings.add(msg);
        }
      }

      if (isString && data[fieldName].length > 1) {
        if (operation === "import") {
          validatorWarnings.add(
            intl.formatMessage(
              {
                id: "validation.messages.multipleValuesWarning",
                defaultMessage:
                  "Multiple values found in field <b>{fieldName}</b> at id {dataId} on table <b>{tableName}</b>. Only the first value {firstValue} will be considered",
              },
              {
                fieldName,
                dataId: data["@id"],
                tableName,
                firstValue: data[fieldName][0],
                b: (str) => `<b>${str}</b>`,
              }
            )
          );
        } else {
          validatorErrors.add(
            intl.formatMessage(
              {
                id: "validation.messages.multipleValues",
                defaultMessage:
                  "Multiple values found in field <b>{fieldName}</b> at id {dataId} on table <b>{tableName}</b>.",
              },
              {
                fieldName,
                dataId: data["@id"],
                tableName,
                b: (str) => `<b>${str}</b>`,
              }
            )
          );
        }
        data[fieldName] = [data[fieldName][0]];
      }

      const linkedTable = field.link?.table.className;
      const linkedIds: string[] = [];

      if (predefinedCodeLists.includes(linkedTable)) {
        const codeList = await getCodeListByTableName(linkedTable);
        if (codeList) {
          codeList.forEach((item) => {
            linkedIds.push(item["@id"]);
          });
        }
      }

      for (const linkedData of tableData) {
        if (validateTypeProp(linkedData, intl)) return;
        const linkedTableName = linkedData["@type"].split(":")[1];
        if (linkedTableName === linkedTable) {
          linkedIds.push(linkedData["@id"]);
        }
      }

      data[fieldName].forEach((item) => {
        if (!linkedIds.includes(item)) {
          validatorWarnings.add(
            intl.formatMessage(
              {
                id: "validation.messages.linkedFieldNotFound",
                defaultMessage:
                  "{tableName} <b>{name}</b> linked on {fieldName} to item <b>{item}</b> that does not exist in the {linkedTable} table",
              },
              {
                tableName,
                name: data["org:hasLegalName"] || data["hasLegalName"] || data["hasName"],
                fieldName,
                item,
                linkedTable,
                b: (str: string) => `<b>${str}</b>`,
              }
            ) as string
          );
        }
      });

      if ((isString || field.representedType === "string") && operation === "export") {
        data[fieldName] = data[fieldName].length > 0 ? data[fieldName][0] : "";
      }
    });
  }
}

function removeEmptyRows(tableData: TableInterface[]) {
  return tableData.filter((item) => item["@id"] && item["@id"].length > 0);
}

function validateIfIdIsValidUrl(
  tableData: TableInterface[],
  operation: Operation,
  intl: IntlShape
) {
  tableData.map((item) => {
    let tableName;
    try {
      tableName = item["@type"].split(":")[1];
    } catch (error) {
      validatorErrors.add(
        intl.formatMessage({
          id: "validation.messages.missingTypeProperty",
          defaultMessage: "<b>@type</b> must be present in the data",
        })
      );
    }

    const id = item["@id"];
    try {
      new URL(item["@id"]);
    } catch (error) {
      if (operation === "import") {
        validatorWarnings.add(
          intl.formatMessage(
            {
              id: "validation.messages.invalidIdFormat",
              defaultMessage:
                "Invalid URL format: <b>{id}</b> for <b>@id</b> on table <b>{tableName}</b>",
            },
            {
              id,
              tableName,
              b: (str) => `<b>${str}</b>`,
            }
          )
        );
        return;
      }
      validatorErrors.add(
        intl.formatMessage(
          {
            id: "validation.messages.invalidIdFormat",
            defaultMessage:
              "Invalid URL format: <b>{id}</b> for <b>@id</b> on table <b>{tableName}</b>",
          },
          {
            id,
            tableName,
            b: (str) => `<b>${str}</b>`,
          }
        )
      );
      return;
    }
  });
}

function validateIfEmptyFile(tableData: TableInterface[], intl: IntlShape) {
  if (!Array.isArray(tableData) || tableData.length === 0) {
    validatorErrors.add(
      intl.formatMessage({
        id: "validation.messages.dataIsEmptyOrNotArray",
        defaultMessage: "Table data is empty or not an array",
      })
    );
  }
}

function isFieldValueNullOrEmpty(value: any) {
  if (typeof value === "string") {
    return value.trim().length === 0;
  }
  if (Array.isArray(value)) {
    return value.length === 0;
  }
  if (typeof value === "object") {
    return Object.keys(value).length === 0;
  }
  return value === null || value === undefined;
}
