/* global document fetch setTimeout clearTimeout URL Blob */
import * as jsonld from "jsonld";
import { Options } from "jsonld";
import { contextUrl } from "../domain/models";

export function downloadJSONLD(data: any, filename: string): void {
  const jsonLDString = JSON.stringify(data, null, 2);
  const blob = new Blob([jsonLDString], { type: "application/ld+json" });
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.setAttribute("download", filename);
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

// Custom document loader
const customLoader: Options.DocLoader["documentLoader"] = async (url: string) => {
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

  try {
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

export async function parseJsonLd(jsonLdData: any) {
  try {
    // Expand JSON-LD if needed
    const expandedData = Array.isArray(jsonLdData)
      ? await jsonld.expand(jsonLdData, {
          documentLoader: customLoader,
        })
      : jsonLdData;

    // Compact JSON-LD using the CIDS context
    const compactedData = await jsonld.compact(
      expandedData,
      {
        "@context": contextUrl,
      },
      {
        documentLoader: customLoader,
      }
    );

    const instances = (compactedData["@graph"] as any[]) || [];

    return instances;
  } catch (error: any) {
    throw new Error(`Error parsing JSON-LD: ${error.message}`);
  }
}
