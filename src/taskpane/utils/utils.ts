/* global document fetch setTimeout clearTimeout URL Blob */
import * as jsonld from "jsonld";
import { Options } from "jsonld";
import { getContext } from "../domain/fetchServer/getContext";
import { contextUrl, map, mapSFFModel } from "../domain/models";

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
    // The object already uses a good context; return it as a single-element array.
    if (Array.isArray(obj["@type"])) {
      // eslint-disable-next-line no-param-reassign
      obj["@type"] = findFirstRecognizedType(obj["@type"]);
    }
    return [obj];
  } else {
    // Otherwise, process it:
    const expanded = await jsonld.expand(obj, { documentLoader: customLoader });

    // Fetch both contexts dynamically
    const [defaultContextData, sffContextData] = await Promise.all([
      getContext(contextUrl[0]),
      getContext(contextUrl[1]),
    ]);

    const mergedContext = [
      (defaultContextData as any)["@context"],
      (sffContextData as any)["@context"],
    ] as unknown as jsonld.ContextDefinition;

    const compacted = await jsonld.compact(expanded, mergedContext, {
      documentLoader: customLoader,
    });

    // The compacted document might contain an @graph.
    let instances = (compacted["@graph"] as any[]) || [compacted];

    // Apply URL replacement to each instance.
    instances = instances.map((instance) => replaceOldUrls(instance));

    // check if @type is an array if yes we find first recognized type
    instances.forEach((instance) => {
      if (Array.isArray(instance["@type"])) {
        // eslint-disable-next-line no-param-reassign
        instance["@type"] = findFirstRecognizedType(instance["@type"]);
      }
    });

    return instances;
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
    results = removeCidsPrefix(results);
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
  for (const key in obj) {
    if (["@context", "@id", "@type"].includes(key)) {
      continue;
    }
    if (!key.startsWith("cids:")) {
      const cidsKey = "cids:" + key;
      // eslint-disable-next-line no-prototype-builtins
      if (obj.hasOwnProperty(cidsKey)) {
        // eslint-disable-next-line no-param-reassign
        delete obj[cidsKey];
      }
    }
    if (typeof obj[key] === "object" && obj[key] !== null && !Array.isArray(obj[key])) {
      // eslint-disable-next-line no-param-reassign
      obj[key] = cleanupDuplicates(obj[key]);
    }
  }
  return obj;
}

// recursively remove cids: prefix from keys
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
