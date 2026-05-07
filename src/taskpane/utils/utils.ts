/* global document fetch setTimeout clearTimeout URL Blob FileReader AbortController */
import * as jsonld from "jsonld";
import { Options } from "jsonld";
import { IntlShape, MessageDescriptor } from "react-intl";
import { getContext } from "../domain/fetchServer/getContext";
import { getUnitOptions } from "../domain/fetchServer/getUnitsOfMeasure";
import { contextUrl, map, mapSFFModel } from "../domain/models";

/**
 * @param obj - The object or array to process
 * @returns The object with updated property names
 **/
export function convertNumericalValueToHasNumericalValue(obj: any): any {
  if (Array.isArray(obj)) return obj.map(convertNumericalValueToHasNumericalValue);
  if (obj && typeof obj === "object") {
    const clone: any = { ...obj };
    if (
      Object.prototype.hasOwnProperty.call(clone, "i72:numerical_value") ||
      Object.prototype.hasOwnProperty.call(clone, "numerical_value")
    ) {
      const value = clone["i72:numerical_value"] ?? clone["numerical_value"];
      clone["i72:hasNumericalValue"] = value;
      delete clone["i72:numerical_value"];
      delete clone["numerical_value"];
    }
    for (const key of Object.keys(clone)) {
      clone[key] = convertNumericalValueToHasNumericalValue(clone[key]);
    }
    return clone;
  }
  return obj;
}

/**
 * @param obj - The object or array to process
 * @param validUnitIds - Array of valid unit IDs from getUnitOptions()
 * @returns Object with converted data and conversion flag
 */
export async function convertUnknownUnitToDescription(
  obj: any,
  validUnitIds?: string[]
): Promise<{ data: any; converted: boolean }> {
  let converted = false;
  let unitIds = validUnitIds;
  if (!unitIds) {
    try {
      const unitOptions: Array<{ id: string }> = await getUnitOptions();
      unitIds = unitOptions.map((option: { id: string }) => option.id);
    } catch (_e) {
      return { data: obj, converted: false };
    }
  }
  async function processItem(item: any): Promise<any> {
    if (Array.isArray(item)) return Promise.all(item.map(processItem));
    if (item && typeof item === "object") {
      const isIndicator =
        item["@type"] &&
        ((typeof item["@type"] === "string" &&
          (item["@type"] === "cids:Indicator" || item["@type"] === "Indicator")) ||
          (Array.isArray(item["@type"]) &&
            item["@type"].some((t: string) => t === "cids:Indicator" || t === "Indicator")));
      let working: any = item;
      if (isIndicator) {
        const unitOfMeasure = item["i72:unit_of_measure"];
        const unitDescription = item["unitDescription"];
        if (
          unitOfMeasure &&
          unitIds &&
          !unitIds.includes(unitOfMeasure) &&
          (!unitDescription || unitDescription === "")
        ) {
          working = { ...item, unitDescription: unitOfMeasure };
          converted = true;
        }
      }
      const result: any = { ...working };
      for (const key of Object.keys(result)) {
        result[key] = await processItem(result[key]);
      }
      return result;
    }
    return item;
  }
  const processedData = await processItem(obj);
  return { data: processedData, converted };
}

/**
 * Checks if a string starts with a BOM (Byte Order Mark).
 */
export function hasBOM(text: string): boolean {
  return text.charCodeAt(0) === 0xfeff;
}

/**
 * Removes BOM (Byte Order Mark) from the start of a string if present.
 */
export function stripBOM(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

/**
 * Converts an old ic:Address object to the new schema:PostalAddress/cids:Address format.
 */
export function convertIcAddressToPostalAddress(obj: any): any {
  if (!obj || typeof obj !== "object") return obj;
  const isAddressType =
    (typeof obj["@type"] === "string" && obj["@type"].toLowerCase().includes("address")) ||
    (Array.isArray(obj["@type"]) &&
      obj["@type"].some((t: string) => t.toLowerCase().includes("address")));
  const hasOldFields =
    obj["ic:hasStreet"] ||
    obj["ic:hasStreetNumber"] ||
    obj["ic:hasStreetType"] ||
    obj["ic:hasStreetDirection"];
  if (isAddressType) {
    if (hasOldFields) {
      const streetNumber = obj["ic:hasStreetNumber"] || "";
      const street = obj["ic:hasStreet"] || "";
      const streetType = obj["ic:hasStreetType"] ? obj["ic:hasStreetType"].replace(/^ic:/, "") : "";
      const streetDirection = obj["ic:hasStreetDirection"]
        ? obj["ic:hasStreetDirection"].replace(/^ic:/, "")
        : "";
      const streetParts = [streetNumber, street, streetType, streetDirection].filter(Boolean);
      const streetAddress = streetParts.join(" ").trim();
      const newAddress: any = { streetAddress };
      const mappings: Record<string, string> = {
        "ic:hasUnitNumber": "extendedAddress",
        "ic:hasCity": "addressLocality",
        "ic:hasState": "addressRegion",
        "ic:hasPostalCode": "postalCode",
        "ic:hasCountry": "addressCountry",
        "ic:hasPostOfficeBoxNumber": "postOfficeBoxNumber",
      };
      for (const [oldKey, newKey] of Object.entries(mappings)) {
        if (obj[oldKey]) newAddress[newKey] = obj[oldKey];
      }
      if (obj["@id"]) newAddress["@id"] = obj["@id"];
      if (obj["@type"]) {
        if (typeof obj["@type"] === "string") {
          newAddress["@type"] = obj["@type"].replace(/^ic:Address$/i, "cids:Address");
        } else if (Array.isArray(obj["@type"])) {
          newAddress["@type"] = obj["@type"].map((type: string) =>
            type.replace(/^ic:Address$/i, "cids:Address")
          );
        } else {
          newAddress["@type"] = obj["@type"];
        }
      }
      return newAddress;
    }
    const typeAddress = { ...obj };
    if (obj["@type"]) {
      if (typeof obj["@type"] === "string") {
        typeAddress["@type"] = obj["@type"].replace(/^ic:Address$/i, "cids:Address");
      } else if (Array.isArray(obj["@type"])) {
        typeAddress["@type"] = obj["@type"].map((t: string) =>
          t.replace(/^ic:Address$/i, "cids:Address")
        );
      }
    }
    return typeAddress;
  }
  const clone: any = { ...obj };
  for (const key of Object.keys(obj)) {
    if (obj[key] && typeof obj[key] === "object") {
      clone[key] = convertIcAddressToPostalAddress(obj[key]);
    }
  }
  return clone;
}

/**
 * Converts ic:hasAddress property to hasAddress in Organization objects and other objects.
 */
export function convertIcHasAddressToHasAddress(obj: any): any {
  if (!obj || typeof obj !== "object") return obj;
  if (Array.isArray(obj)) {
    return obj.map(convertIcHasAddressToHasAddress);
  }
  const newObj = { ...obj };
  if (newObj["ic:hasAddress"]) {
    newObj["hasAddress"] = newObj["ic:hasAddress"];
    delete newObj["ic:hasAddress"];
  }
  for (const key of Object.keys(newObj)) {
    if (newObj[key] && typeof newObj[key] === "object") {
      newObj[key] = convertIcHasAddressToHasAddress(newObj[key]);
    }
  }
  return newObj;
}

/**
 * Converts forFunderId property to forOrganization in FundingStatus objects.
 */
export function convertForFunderIdToForOrganization(obj: any): any {
  if (!obj || typeof obj !== "object") return obj;
  if (Array.isArray(obj)) {
    return obj.map(convertForFunderIdToForOrganization);
  }
  const newObj = { ...obj };
  const isFundingStatus =
    newObj["@type"] &&
    ((typeof newObj["@type"] === "string" &&
      (newObj["@type"] === "cids:FundingStatus" ||
        newObj["@type"] === "FundingStatus" ||
        newObj["@type"] === "sff:FundingStatus")) ||
      (Array.isArray(newObj["@type"]) &&
        newObj["@type"].some(
          (t: string) =>
            t === "cids:FundingStatus" || t === "FundingStatus" || t === "sff:FundingStatus"
        )));
  if (isFundingStatus && newObj["forFunderId"]) {
    if (!newObj["forOrganization"]) {
      newObj["forOrganization"] = newObj["forFunderId"];
    }
    delete newObj["forFunderId"];
  }
  for (const key of Object.keys(newObj)) {
    if (newObj[key] && typeof newObj[key] === "object") {
      newObj[key] = convertForFunderIdToForOrganization(newObj[key]);
    }
  }
  return newObj;
}

/**
 * Converts 'identifier' field to 'org:hasIdentifier' in OrganizationID objects for backward compatibility.
 */
export function convertOrganizationIDFields(obj: any): any {
  if (Array.isArray(obj)) {
    return obj.map(convertOrganizationIDFields);
  } else if (obj && typeof obj === "object") {
    const typeVal = obj["@type"];
    const isOrganizationID =
      typeVal &&
      ((typeof typeVal === "string" &&
        (typeVal === "org:OrganizationID" ||
          typeVal === "sff:OrganizationID" ||
          typeVal === "OrganizationID")) ||
        (Array.isArray(typeVal) &&
          typeVal.some(
            (t: string) =>
              t === "org:OrganizationID" || t === "sff:OrganizationID" || t === "OrganizationID"
          )));
    if (isOrganizationID) {
      if (typeof typeVal === "string" && typeVal !== "org:OrganizationID") {
        obj["@type"] = "org:OrganizationID";
      } else if (Array.isArray(typeVal)) {
        obj["@type"] = typeVal.map((t: string) =>
          t === "sff:OrganizationID" || t === "OrganizationID" ? "org:OrganizationID" : t
        );
      }
      if (
        Object.prototype.hasOwnProperty.call(obj, "identifier") &&
        !Object.prototype.hasOwnProperty.call(obj, "org:hasIdentifier") &&
        !Object.prototype.hasOwnProperty.call(obj, "hasIdentifier")
      ) {
        obj["org:hasIdentifier"] = obj["identifier"];
        delete obj["identifier"];
      }
      if (
        Object.prototype.hasOwnProperty.call(obj, "issuedBy") &&
        !Object.prototype.hasOwnProperty.call(obj, "org:issuedBy")
      ) {
        obj["org:issuedBy"] = obj["issuedBy"];
        delete obj["issuedBy"];
      }
    }
    for (const key of Object.keys(obj)) {
      obj[key] = convertOrganizationIDFields(obj[key]);
    }
    return obj;
  }
  return obj;
}

// Accept both describesPopulation (cids alias) and i72:cardinality_of (ontology);
export function harmonizeCardinalityProperty(data: any): any {
  const sync = (input: any): any => {
    if (!input || typeof input !== "object") return input;
    const cloned: any = Array.isArray(input) ? [...input] : { ...input };
    const hasCard = Object.prototype.hasOwnProperty.call(cloned, "i72:cardinality_of");
    const hasDesc = Object.prototype.hasOwnProperty.call(cloned, "describesPopulation");
    if (hasCard && !hasDesc) {
      cloned.describesPopulation = cloned["i72:cardinality_of"];
      delete cloned["i72:cardinality_of"];
    } else if (hasCard && hasDesc) {
      delete cloned["i72:cardinality_of"];
    }
    return cloned;
  };
  if (Array.isArray(data)) return data.map(sync);
  if (data && typeof data === "object") return sync(data);
  return data;
}

/**
 * Formats an intl message to a plain string, handling ReactNode[] from rich text formatting.
 */
export function formatMessageToString(
  intl: IntlShape,
  descriptor: MessageDescriptor,
  values?: Record<string, any>
): string {
  const message = intl.formatMessage(descriptor, values);
  if (Array.isArray(message)) {
    return message.map((part) => (typeof part === "string" ? part : "")).join("");
  }
  return message as string;
}

/**
 * Handles the change event of a file input element.
 */
export const handleFileChange = async (
  event: any,
  onSuccess: (data: any) => void,
  onError: (error: any) => void,
  intl: IntlShape
): Promise<void> => {
  const file = event.target.files[0];
  if (file && (file.name.endsWith(".jsonld") || file.name.endsWith(".json"))) {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const target = e?.target as FileReader | null;
        const raw = target && typeof target.result === "string" ? target.result : "";
        const data = JSON.parse(raw as any);
        onSuccess(data);
      } catch (error) {
        onError(
          new Error(
            intl.formatMessage({
              id: "import.messages.error.notValidJson",
              defaultMessage: "File is not a valid JSON/JSON-LD file.",
            })
          )
        );
      }
    };
    reader.readAsText(file);
  } else {
    onError(
      new Error(
        intl.formatMessage({
          id: "import.messages.error.notJson",
          defaultMessage: "File is not a JSON/JSON-LD file.",
        })
      )
    );
  }
};

/**
 * Downloads a JSON-LD file.
 */
export function downloadJSONLD(data: any, filename: string): void {
  const jsonLDString = JSON.stringify(data, null, 2);
  const isSafari = /^((?!chrome|android).)*safari/i.test(navigator.userAgent);

  if (isSafari) {
    const dataUri = "data:application/ld+json;charset=utf-8," + encodeURIComponent(jsonLDString);
    const link = document.createElement("a");
    link.href = dataUri;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  } else {
    const blob = new Blob([jsonLDString], { type: "application/ld+json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.target = "_blank";
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    setTimeout(() => {
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }, 1000);
  }
}

/**
 * Executes tasks in batches.
 */
export async function executeInBatches<T>(
  items: T[],
  task: (batch: T[]) => Promise<void>,
  batchSize: number = 50
): Promise<void> {
  for (let i = 0; i < items.length; i += batchSize) {
    const batch = items.slice(i, i + batchSize);
    await task(batch);
  }
}

const trustedDomains = [
  "ontology.commonapproach.org",
  "sparql.cwrc.ca",
  "www.w3.org",
  "xmlns.com",
  "www.opengis.net",
  "schema.org",
  "ontology.eil.utoronto.ca",
];

const customLoader: Options.DocLoader["documentLoader"] = async (url: string) => {
  try {
    if (contextUrl.includes(url)) {
      const context = await getContext(url);
      return {
        contextUrl: undefined,
        documentUrl: url,
        document: context,
      };
    }
    if (url.startsWith("http://")) {
      // eslint-disable-next-line no-param-reassign
      url = url.replace("http://", "https://");
    }
    const urlDomain = new URL(url).hostname;
    if (!trustedDomains.some((trustedDomain) => urlDomain.endsWith(trustedDomain))) {
      throw new Error(`URL not trusted: ${url}`);
    }
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 5000);
    const response = await fetch(url, { signal: controller.signal });
    clearTimeout(timeoutId);
    if (!response.ok) {
      throw new Error(`Failed to load context from URL: ${url} (Status: ${response.status})`);
    }
    const document = await response.json();
    return {
      contextUrl: undefined,
      documentUrl: url,
      document,
    };
  } catch (error: any) {
    if (error.name === "AbortError") {
      throw new Error(`Request timed out while trying to load context from URL: ${url}`);
    } else if (error.message.includes("Failed to fetch")) {
      throw new Error(`CORS issue or network error while trying to load context from URL: ${url}`);
    } else {
      throw new Error(`Error loading context from URL: ${url} (${error.message})`);
    }
  }
};

const goodContexts = [
  "https://ontology.commonapproach.org/contexts/cidsContext.jsonld",
  "https://ontology.commonapproach.org/contexts/sffContext.jsonld",
  "https://ontology.commonapproach.org/cids.jsonld",
  "https://ontology.commonapproach.org/sff-1.0.jsonld",
  "http://ontology.eil.utoronto.ca/cids/contexts/cidsContext.json",
  "https://ontology.commonapproach.org/contexts/cidsContext.json",
];

const urlsToReplace = [
  "http://ontology.eil.utoronto.ca/cids/cids#",
  "http://ontology.commonapproach.org/owl/cids_v2.1.owl/cids#",
  "http://ontology.commonapproach.org/tove/organization#",
  "http://ontology.commonapproach.org/ISO21972/iso21972#",
  "https://www.w3.org/Submission/prov-json/schema#",
  "http://ontology.commonapproach.org/tove/icontact#",
];

async function processJsonLdObject(obj: any): Promise<any[]> {
  let alreadyGood = false;
  if (obj["@context"]) {
    if (typeof obj["@context"] === "string") {
      alreadyGood = goodContexts.includes(obj["@context"]);
    } else if (Array.isArray(obj["@context"])) {
      alreadyGood = obj["@context"].some((c: string) => goodContexts.includes(c));
    }
  }
  if (alreadyGood) {
    const base = { ...obj };
    if (Array.isArray(base["@type"])) {
      base["@type"] = findFirstRecognizedType(base["@type"]);
    }
    return [base];
  } else {
    const expanded = await jsonld.expand(obj, { documentLoader: customLoader });
    const [defaultContextData, sffContextData] = await Promise.all([
      getContext(contextUrl[0]) as Promise<any>,
      getContext(contextUrl[1]) as Promise<any>,
    ]);
    const mergedContext = [
      defaultContextData["@context"],
      sffContextData["@context"],
    ] as unknown as jsonld.ContextDefinition;
    const compacted = await jsonld.compact(expanded, mergedContext, {
      documentLoader: customLoader,
    });
    let instances = (compacted["@graph"] as any[]) || [compacted];
    instances = instances.map((instance) => replaceOldUrls(instance));
    return instances.map((instance) => {
      if (Array.isArray(instance["@type"])) {
        return { ...instance, "@type": findFirstRecognizedType(instance["@type"]) };
      }
      return instance;
    });
  }
}

export async function parseJsonLd(jsonLdData: any[]): Promise<any[]> {
  const processedInstances: any[] = [];
  for (const obj of jsonLdData) {
    let results = await processJsonLdObject(obj);
    results = cleanupDuplicates(results);
    results = removeCidsPrefix(results);
    processedInstances.push(...results);
  }
  return processedInstances;
}

function replaceOldUrls(input: any): any {
  if (typeof input === "string") {
    for (const url of urlsToReplace) {
      if (input.startsWith(url)) {
        return input.replace(url, "cids:");
      }
    }
    return input;
  } else if (Array.isArray(input)) {
    return input.map((item) => replaceOldUrls(item));
  } else if (input !== null && typeof input === "object") {
    // eslint-disable-next-line no-prototype-builtins
    if (input.hasOwnProperty("@value")) {
      return replaceOldUrls(input["@value"]);
    }
    const newObj: any = {};
    for (const [key, value] of Object.entries(input)) {
      let newKey = key;
      if (key.startsWith("http://") || key.startsWith("https://")) {
        const matchedUrl = urlsToReplace.find((url) => key.startsWith(url));
        if (matchedUrl) {
          const parts = key.split("#");
          if (parts.length > 1 && parts[1].length > 0) {
            newKey = "cids:" + parts[1];
          }
        }
        newObj[newKey] = replaceOldUrls(value);
      } else {
        newObj[key] = replaceOldUrls(value);
      }
    }
    for (const key in newObj) {
      // eslint-disable-next-line no-prototype-builtins
      if (!key.includes(":") && newObj.hasOwnProperty("cids:" + key)) {
        delete newObj["cids:" + key];
      }
    }
    return newObj;
  }
  return input;
}

function cleanupDuplicates(obj: any): any {
  if (!obj || typeof obj !== "object" || Array.isArray(obj)) return obj;
  const clone: any = { ...obj };
  for (const key of Object.keys(obj)) {
    if (["@context", "@id", "@type"].includes(key)) continue;
    if (!key.startsWith("cids:")) {
      const cidsKey = "cids:" + key;
      // eslint-disable-next-line no-prototype-builtins
      if (Object.prototype.hasOwnProperty.call(clone, cidsKey)) delete clone[cidsKey];
    }
    if (typeof obj[key] === "object" && obj[key] !== null && !Array.isArray(obj[key])) {
      clone[key] = cleanupDuplicates(obj[key]);
    }
  }
  return clone;
}

function removeCidsPrefix(obj: any): any {
  if (Array.isArray(obj)) {
    return obj.map((item) => removeCidsPrefix(item));
  } else if (obj !== null && typeof obj === "object") {
    const newObj: any = {};
    for (const [key, value] of Object.entries(obj)) {
      let newKey = key;
      if (key.startsWith("cids:")) {
        newKey = key.substring(5);
      }
      newObj[newKey] = removeCidsPrefix(value);
    }
    return newObj;
  }
  return obj;
}

function extractClassName(type: string): string {
  if (type.includes(":")) {
    return type.split(":").pop() || "";
  }
  return type.split(/[/#]/).pop() || "";
}

function findFirstRecognizedType(types: string | string[]): string {
  if (!Array.isArray(types)) {
    return types;
  }
  const recognizedTypes = new Set([...Object.keys(map), ...Object.keys(mapSFFModel)]);
  for (const type of types) {
    const className = extractClassName(type);
    if (recognizedTypes.has(className)) {
      return type;
    }
  }
  return types[0];
}