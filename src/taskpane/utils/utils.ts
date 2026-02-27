/* global document fetch setTimeout clearTimeout URL Blob FileReader AbortController */
import * as jsonld from "jsonld";
import { Options } from "jsonld";
import { IntlShape, MessageDescriptor } from "react-intl";
import { getContext } from "../domain/fetchServer/getContext";
import { getUnitOptions } from "../domain/fetchServer/getUnitsOfMeasure";
import { contextUrl, map, mapSFFModel } from "../domain/models";

/**
 * Recursively converts any i72:numerical_value or numerical_value (with or without prefix)
 * to i72:hasNumericalValue in imported data, for backward compatibility.
 * @param obj - The object or array to process
 * @returns The object with updated property names
 */
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
 * Converts unknown unit_of_measure values to unitDescription for backward compatibility.
 * This is used when importing Indicator objects with unknown unit values that are not
 * in our unit of measure list. The unknown value is copied to unitDescription field
 * so users can see and fix it manually, but only if unitDescription is not already set.
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
 * @param text - The input string
 * @returns true if BOM is present, false otherwise
 */
export function hasBOM(text: string): boolean {
  return text.charCodeAt(0) === 0xfeff;
}
/**
 * Removes BOM (Byte Order Mark) from the start of a string if present.
 * @param text - The input string
 * @returns The string without BOM
 */
export function stripBOM(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}
/**
 * Converts an old ic:Address object to the new schema:PostalAddress/cids:Address format.
 * Returns a new object with PostalAddress fields, or the original if not an old address.
 * If the object is not an Address, recursively checks its fields for address objects.
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
 * This handles backward compatibility for the property name change.
 */
export function convertIcHasAddressToHasAddress(obj: any): any {
  if (!obj || typeof obj !== "object") return obj;

  // Handle arrays
  if (Array.isArray(obj)) {
    return obj.map(convertIcHasAddressToHasAddress);
  }

  // Create a new object to avoid mutating the original
  const newObj = { ...obj };

  // Convert ic:hasAddress to hasAddress if it exists
  if (newObj["ic:hasAddress"]) {
    newObj["hasAddress"] = newObj["ic:hasAddress"];
    delete newObj["ic:hasAddress"];
  }

  // Recursively process nested objects
  for (const key of Object.keys(newObj)) {
    if (newObj[key] && typeof newObj[key] === "object") {
      newObj[key] = convertIcHasAddressToHasAddress(newObj[key]);
    }
  }

  return newObj;
}

/**
 * Converts forFunderId property to forOrganization in FundingStatus objects.
 * This handles backward compatibility for the property name change.
 */
export function convertForFunderIdToForOrganization(obj: any): any {
  if (!obj || typeof obj !== "object") return obj;

  // Handle arrays
  if (Array.isArray(obj)) {
    return obj.map(convertForFunderIdToForOrganization);
  }

  // Create a new object to avoid mutating the original
  const newObj = { ...obj };

  // Check if this is a FundingStatus object (by @type)
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

  // Convert forFunderId to forOrganization if it exists and we are in a FundingStatus object
  if (isFundingStatus && newObj["forFunderId"]) {
    if (!newObj["forOrganization"]) {
      newObj["forOrganization"] = newObj["forFunderId"];
    }
    delete newObj["forFunderId"];
  }

  // Recursively process nested objects
  for (const key of Object.keys(newObj)) {
    if (newObj[key] && typeof newObj[key] === "object") {
      newObj[key] = convertForFunderIdToForOrganization(newObj[key]);
    }
  }

  return newObj;
}

// Accept both describesPopulation (cids alias) and i72:cardinality_of (ontology);
// prefer describesPopulation in stored/processed objects and remove i72:cardinality_of after aliasing.
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
 * Handles the change event of a file input element.
 * Reads the selected file, checks if it has a ".jsonld" extension,
 * and if so, reads the file content as text and parses it as JSON.
 * Finally, calls the `onSuccess` callback function with the parsed JSON data.
 * @param event - The event object triggered by the file input element.
 * @param onSuccess - A callback function that will be called with the parsed JSON data.
 * @param onError - A callback function that will be called if an error occurs.
 * @returns Promise<void>
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
 * Downloads a JSON-LD file by converting the data into a JSON string,
 * creating a Blob object with the JSON string, and generating a download link for the Blob object.
 * When the link is clicked, the file is downloaded.
 *
 * @param data - The data to be downloaded as a JSON-LD file.
 * @param filename - The name of the downloaded file.
 * @returns void
 */
export function downloadJSONLD(data: any, filename: string): void {
  const jsonLDString = JSON.stringify(data, null, 2);
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
  }, 1000); // Wait for 1 second before removing the link and revoking the URL
}

/**
 * Executes tasks in batches, where each task operates on a batch of items.
 * @param items - The array of items to be processed.
 * @param task - A function that processes a batch of items and returns a Promise.
 * @param batchSize - The number of items to process in each batch. (default: 50)
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

/**
 * Custom document loader that enforces HTTPS, checks for trusted domains, and fetches JSON-LD context documents.
 *
 * @param url - The URL of the context document to load.
 * @returns A promise that resolves to an object containing the context document.
 * @throws Will throw an error if the URL is not trusted, if the request times out, or if there is a network/CORS issue.
 */
const customLoader: Options.DocLoader["documentLoader"] = async (url: string) => {
  try {
    // Get default context
    if (contextUrl.includes(url)) {
      const context = await getContext(url);

      return {
        contextUrl: undefined,
        documentUrl: url,
        document: context,
      };
    }

    // Enforce HTTPS by rewriting the URL
    if (url.startsWith("http://")) {
      // eslint-disable-next-line no-param-reassign
      url = url.replace("http://", "https://");
    }

    // Check if the URL is in the trusted list
    const urlDomain = new URL(url).hostname;
    if (!trustedDomains.some((trustedDomain) => urlDomain.endsWith(trustedDomain))) {
      throw new Error(`URL not trusted: ${url}`);
    }
    // Fetch the context document using HTTPS
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 5000); // Set a timeout of 5 seconds

    const response = await fetch(url, { signal: controller.signal }); // Use AbortController's signal
    clearTimeout(timeoutId);
    if (!response.ok) {
      throw new Error(`Failed to load context from URL: ${url} (Status: ${response.status})`);
    }
    const document = await response.json();

    // Return the fetched document
    return {
      contextUrl: undefined, // No additional context
      documentUrl: url, // The URL of the document
      document, // The resolved JSON-LD context
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

// List of good context URLs.
const goodContexts = [
  "https://ontology.commonapproach.org/contexts/cidsContext.jsonld", // Base context
  "https://ontology.commonapproach.org/contexts/sffContext.jsonld", // Extended context for SFF module
  "https://ontology.commonapproach.org/cids.jsonld",
  "https://ontology.commonapproach.org/sff-1.0.jsonld",
  "http://ontology.eil.utoronto.ca/cids/contexts/cidsContext.json", // try to keep compatibility with old CIDS ontology
  "https://ontology.commonapproach.org/contexts/cidsContext.json", // try to keep compatibility with old CIDS ontology
];

// Replace URL list to try to keep minimal compatibility with the old CIDS ontology.
const urlsToReplace = [
  "http://ontology.eil.utoronto.ca/cids/cids#",
  "http://ontology.commonapproach.org/owl/cids_v2.1.owl/cids#",
  "http://ontology.commonapproach.org/tove/organization#",
  "http://ontology.commonapproach.org/ISO21972/iso21972#",
  "https://www.w3.org/Submission/prov-json/schema#",
  "http://ontology.commonapproach.org/tove/icontact#",
];

/**
 * Processes a single JSON-LD object.
 * If the object already uses one of the good contexts, it is returned as-is.
 * Otherwise, it is expanded, compacted with the merged context, and then processed
 * to replace legacy URL portions.
 */
async function processJsonLdObject(obj: any): Promise<any[]> {
  // Check if the object's @context is already one of the good ones.
  let alreadyGood = false;
  if (obj["@context"]) {
    if (typeof obj["@context"] === "string") {
      alreadyGood = goodContexts.includes(obj["@context"]);
    } else if (Array.isArray(obj["@context"])) {
      alreadyGood = obj["@context"].some((c: string) => goodContexts.includes(c));
    }
  }

  if (alreadyGood) {
    // clone to avoid param mutation warnings
    const base = { ...obj };
    if (Array.isArray(base["@type"])) {
      base["@type"] = findFirstRecognizedType(base["@type"]);
    }
    return [base];
  } else {
    // Otherwise, process it:
    const expanded = await jsonld.expand(obj, { documentLoader: customLoader });

    // Fetch both contexts dynamically
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

    // The compacted document might contain an @graph.
    let instances = (compacted["@graph"] as any[]) || [compacted];

    // Apply URL replacement to each instance.
    instances = instances.map((instance) => replaceOldUrls(instance));

    // check if @type is an array if yes we find first recognized type
    return instances.map((instance) => {
      if (Array.isArray(instance["@type"])) {
        return { ...instance, "@type": findFirstRecognizedType(instance["@type"]) };
      }
      return instance;
    });
  }
}

/**
 * Iterates through the array of JSON-LD objects and processes each one individually.
 */
export async function parseJsonLd(jsonLdData: any[]): Promise<any[]> {
  const processedInstances: any[] = [];
  for (const obj of jsonLdData) {
    let results = await processJsonLdObject(obj);
    results = cleanupDuplicates(results);
    results = removeAllNamespacePrefixes(results);
    processedInstances.push(...results);
  }
  return processedInstances;
}

/**
 * Recursively replace legacy URLs in strings and object keys.
 * Only acts on keys that are full IRIs (i.e. start with "http://" or "https://").
 * If a key is not a full IRI (already compact), it is left unchanged.
 */
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
    // If this is a "value object", flatten it.
    // eslint-disable-next-line no-prototype-builtins
    if (input.hasOwnProperty("@value")) {
      return replaceOldUrls(input["@value"]);
    }
    const newObj: any = {};
    for (const [key, value] of Object.entries(input)) {
      let newKey = key;
      // Process keys that appear to be full IRIs.
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
    // Clean up duplicate keys if they exist.
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

// recursively remove cids: prefix from keys
function removeAllNamespacePrefixes(obj: any): any {
  if (Array.isArray(obj)) {
    return obj.map((item) => removeAllNamespacePrefixes(item));
  } else if (obj !== null && typeof obj === "object") {
    const newObj: any = {};

    for (const [key, value] of Object.entries(obj)) {
      let newKey = key;

      // Preserve JSON-LD keywords (start with @)
      if (key.startsWith("@")) {
        newObj[key] = removeAllNamespacePrefixes(value);
        continue;
      }

      // Remove any namespace prefix (anything before and including ":")
      if (key.includes(":")) {
        const parts = key.split(":");
        if (parts.length === 2 && parts[0].length > 0 && parts[1].length > 0) {
          // Remove the prefix (e.g., "rdfs:label" => "label")
          newKey = parts[1];
        }
      }

      // Recursively process the value
      newObj[newKey] = removeAllNamespacePrefixes(value);
    }

    return newObj;
  }

  return obj;
}

// Extract class name from any format
function extractClassName(type: string): string {
  // Handle prefixed format (e.g., "cids:MyClass")
  if (type.includes(":")) {
    return type.split(":").pop() || "";
  }
  // Handle URL format (e.g., "http://example.com/MyClass" or "http://example.com/example#MyClass")
  return type.split(/[/#]/).pop() || "";
}

function findFirstRecognizedType(types: string | string[]): string {
  if (!Array.isArray(types)) {
    return types;
  }

  // Create a set of recognized class names from both maps
  const recognizedTypes = new Set([...Object.keys(map), ...Object.keys(mapSFFModel)]);

  // Try to find the first matching type
  for (const type of types) {
    const className = extractClassName(type);
    if (recognizedTypes.has(className)) {
      return type;
    }
  }

  // If no recognized type is found, return the first one
  return types[0];
}

/**
 * Helper to format a message and ensure it returns a string.
 * This handles cases where formatMessage returns ReactNode[] due to rich text formatting.
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
 * Converts 'identifier' field to 'org:hasIdentifier' in OrganizationID objects for backward compatibility.
 * Also normalizes the @type from 'org:OrganizationID' to 'sff:OrganizationID'.
 * @param obj - The object or array to process
 * @returns The object with updated property names
 */
export function convertOrganizationIDFields(obj: any): any {
  if (Array.isArray(obj)) {
    return obj.map(convertOrganizationIDFields);
  } else if (obj && typeof obj === "object") {
    // Check if this is an OrganizationID object
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
  // Normalize @type to org:OrganizationID
  if (typeof typeVal === "string" && typeVal !== "org:OrganizationID") {
    obj["@type"] = "org:OrganizationID";
  } else if (Array.isArray(typeVal)) {
    obj["@type"] = typeVal.map((t: string) =>
      t === "sff:OrganizationID" || t === "OrganizationID" ? "org:OrganizationID" : t
    );
  }

      // Convert 'identifier' to 'org:hasIdentifier' if present
      if (
        Object.prototype.hasOwnProperty.call(obj, "identifier") &&
        !Object.prototype.hasOwnProperty.call(obj, "org:hasIdentifier") &&
        !Object.prototype.hasOwnProperty.call(obj, "hasIdentifier")
      ) {
        obj["org:hasIdentifier"] = obj["identifier"];
        delete obj["identifier"];
      }

      // Convert 'issuedBy' to 'org:issuedBy' if present
      if (
        Object.prototype.hasOwnProperty.call(obj, "issuedBy") &&
        !Object.prototype.hasOwnProperty.call(obj, "org:issuedBy")
      ) {
        obj["org:issuedBy"] = obj["issuedBy"];
        delete obj["issuedBy"];
      }
    }

    // Recursively process all properties
    for (const key of Object.keys(obj)) {
      obj[key] = convertOrganizationIDFields(obj[key]);
    }
    return obj;
  }
  return obj;
}
