import { XMLParser } from "fast-xml-parser";

// ============================================================================
// TYPE DEFINITIONS
// ============================================================================

export interface CodeList {
  "@id": string;
  "@type"?: string;
  hasIdentifier: string;
  hasName: string;
  hasDescription?: string;
}

interface CacheItem {
  data: CodeList[];
  timestamp: number;
  expiresIn: number;
}

// ============================================================================
// CONSTANTS
// ============================================================================

const CACHE_EXPIRATION_TIME = 24 * 60 * 60 * 1000; // 24 hours

/** Primary codelist URLs from commonapproach.org - Updated Dec 2025 */
const CODELIST_URLS = {
  ESDCSector: "https://codelist.commonapproach.org/ESDCSector.ttl",
  PopulationServed: "https://codelist.commonapproach.org/PopulationServed.ttl",
  ProvinceTerritory: "https://codelist.commonapproach.org/ProvinceTerritory.ttl",
  FundingState: "https://codelist.commonapproach.org/FundingState.ttl",
  SDGImpacts: "https://codelist.commonapproach.org/SDGImpacts.ttl",
  OrganizationType: "https://codelist.commonapproach.org/OrgTypeGOC.ttl",
  Locality: "https://codelist.commonapproach.org/LocalityStatsCan.ttl",
  CorporateRegistrar: "https://codelist.commonapproach.org/CanadianCorporateRegistries.ttl",
  IRISImpactCategory: "https://codelist.commonapproach.org/IRISImpactCategory.ttl",
} as const;

/** GitHub fallback URLs for redundancy */
const GITHUB_FALLBACK_URLS: Record<string, string> = {
  [CODELIST_URLS.ESDCSector]:
    "https://raw.githubusercontent.com/commonapproach/CodeLists/main/ESDCSector/ESDCSector.ttl",
  [CODELIST_URLS.PopulationServed]:
    "https://raw.githubusercontent.com/commonapproach/CodeLists/main/PopulationServed/PopulationServed.ttl",
  [CODELIST_URLS.ProvinceTerritory]:
    "https://raw.githubusercontent.com/commonapproach/CodeLists/main/ProvinceTerritory/ProvinceTerritory.ttl",
  [CODELIST_URLS.FundingState]:
    "https://raw.githubusercontent.com/commonapproach/CodeLists/main/FundingState/FundingState.ttl",
  [CODELIST_URLS.SDGImpacts]:
    "https://raw.githubusercontent.com/commonapproach/CodeLists/main/SDGImpacts/SDGImpacts.ttl",
  [CODELIST_URLS.OrganizationType]:
    "https://raw.githubusercontent.com/commonapproach/CodeLists/main/OrgTypeGOC/OrgTypeGOC.ttl",
  [CODELIST_URLS.Locality]:
    "https://raw.githubusercontent.com/commonapproach/CodeLists/main/Locality/LocalityStatsCan.ttl",
  [CODELIST_URLS.CorporateRegistrar]:
    "https://raw.githubusercontent.com/commonapproach/CodeLists/main/CanadianCorporateRegistries.ttl",
  [CODELIST_URLS.IRISImpactCategory]:
    "https://raw.githubusercontent.com/commonapproach/CodeLists/main/IRISImpactCategory.ttl",
};

/** Metadata identifiers to skip during parsing */
const METADATA_IDENTIFIERS = new Set([
  "dataset",
  "IRISImpactCategories",
  "CanadianCorporateRegistries",
  "ESDCSector",
  "PopulationServed",
  "ProvinceTerritory",
  "FundingState",
  "SDGImpacts",
  "OrgTypeGOC",
  "LocalityStatsCan",
]);

/** Keywords that indicate a metadata entry */
const METADATA_KEYWORDS = ["Codelist", "Code List", "Categories", "Registries", "Dataset"];

// ============================================================================
// CACHE MANAGEMENT
// ============================================================================

const inMemoryCache: Record<string, CodeList[]> = {};

function getCachedData(url: string): CodeList[] | null {
  if (inMemoryCache[url]?.length > 0) {
    console.log(`✅ Cache hit (memory): ${url}`);
    return inMemoryCache[url];
  }

  const cachedData = localStorage.getItem(url);
  if (!cachedData) {
    return null;
  }

  try {
    const parsedData = JSON.parse(cachedData) as CacheItem;

    if (!parsedData.data || !parsedData.timestamp || !parsedData.expiresIn) {
      localStorage.removeItem(url);
      return null;
    }

    const isExpired = Date.now() - parsedData.timestamp > parsedData.expiresIn;
    if (isExpired) {
      localStorage.removeItem(url);
      return null;
    }

    console.log(`✅ Cache hit (localStorage): ${url}`);
    inMemoryCache[url] = parsedData.data;
    return parsedData.data;
  } catch (error) {
    localStorage.removeItem(url);
    return null;
  }
}

function setCachedData(url: string, data: CodeList[]): void {
  if (!data || data.length === 0) {
    return;
  }

  inMemoryCache[url] = data;

  const cacheItem: CacheItem = {
    data,
    timestamp: Date.now(),
    expiresIn: CACHE_EXPIRATION_TIME,
  };

  try {
    localStorage.setItem(url, JSON.stringify(cacheItem));
    console.log(`💾 Cached ${data.length} entries for ${url}`);
  } catch (error) {
    console.warn("⚠️ Failed to cache in localStorage:", error);
  }
}

export function clearCodeListCache(tableName?: string): void {
  if (tableName) {
    const urlMap: Record<string, string> = {
      Sector: CODELIST_URLS.ESDCSector,
      PopulationServed: CODELIST_URLS.PopulationServed,
      Locality: CODELIST_URLS.Locality,
      ProvinceTerritory: CODELIST_URLS.ProvinceTerritory,
      FundingState: CODELIST_URLS.FundingState,
      SDGImpacts: CODELIST_URLS.SDGImpacts,
      OrganizationType: CODELIST_URLS.OrganizationType,
      CorporateRegistrar: CODELIST_URLS.CorporateRegistrar,
    };

    const url = urlMap[tableName];
    if (url) {
      delete inMemoryCache[url];
      localStorage.removeItem(url);
      console.log(`🗑️ Cleared cache for ${tableName}`);
    }
  } else {
    // Clear all caches
    Object.keys(inMemoryCache).forEach((key) => delete inMemoryCache[key]);
    Object.keys(localStorage).forEach((key) => {
      if (key.includes("codelist.commonapproach.org")) {
        localStorage.removeItem(key);
      }
    });
    console.log("🗑️ Cleared all codelist caches");
  }
}

// ============================================================================
// XML PARSER (for .owl files - kept for backward compatibility)
// ============================================================================

function parseXmlToCodeList(xmlData: string): CodeList[] {
  const parser = new XMLParser({ ignoreAttributes: false });
  const jsonData = parser.parse(xmlData);

  const codeList: CodeList[] = [];
  const descriptions = jsonData["rdf:RDF"]?.["rdf:Description"] || [];
  let baseIdUrl = "";

  const descArray = Array.isArray(descriptions) ? descriptions : [descriptions];

  for (const desc of descArray) {
    if (desc["vann:preferredNamespacePrefix"]) {
      baseIdUrl = desc["@_rdf:about"]?.replace("#dataset", "") || "";
      continue;
    }

    if (!desc["cids:hasIdentifier"] && !desc["cids:hasName"]) {
      continue;
    }

    const entry: CodeList = {
      "@id": desc["@_rdf:about"]?.includes(baseIdUrl)
        ? desc["@_rdf:about"]
        : baseIdUrl + desc["@_rdf:about"],
      hasIdentifier: desc["cids:hasIdentifier"]?.toString() || "",
      hasName: desc["cids:hasName"]?.["#text"]?.toString() || "",
    };

    if (desc["cids:hasDescription"]?.["#text"]) {
      entry.hasDescription = desc["cids:hasDescription"]["#text"].toString();
    } else if (desc["cids:hasDefinition"]?.["#text"]) {
      entry.hasDescription = desc["cids:hasDefinition"]["#text"].toString();
    }

    codeList.push(entry);
  }

  return codeList;
}

// ============================================================================
// TURTLE PARSER (for .ttl files) - TESTED AND VERIFIED
// ============================================================================

/**
 * Checks if an entry is metadata/header that should be skipped
 */
function isMetadataEntry(identifier: string, name: string): boolean {
  // Check if identifier is in known metadata list
  if (METADATA_IDENTIFIERS.has(identifier)) {
    return true;
  }

  // Check if name contains metadata keywords
  return METADATA_KEYWORDS.some((keyword) => name.includes(keyword));
}

function parseTurtleToCodeList(ttlData: string, sourceUrl: string): CodeList[] {
  const codeList: CodeList[] = [];

  // Extract base URI from @prefix declaration
  let baseUri = "https://codelist.commonapproach.org/codeLists/";

  const baseMatch = ttlData.match(/@base\s+<([^>]+)>/m);
  if (baseMatch) {
    baseUri = baseMatch[1];
  }

  const baseUriMatch = ttlData.match(/@prefix\s*:\s*<([^>]+)>/m);
  if (baseUriMatch) {
    const val = baseUriMatch[1];
    if (val === "#") {
      if (!baseUri.endsWith("#") && !baseUri.endsWith("/")) {
        baseUri += "#";
      }
    } else {
      baseUri = val;
    }
  }

  // 🔥 Parse ALL prefix declarations into a map for URI expansion
  const prefixMap: Record<string, string> = {};
  const prefixRegex = /@prefix\s+([a-zA-Z0-9_-]*)\s*:\s*<([^>]+)>\s*\./g;
  let prefixMatch;
  while ((prefixMatch = prefixRegex.exec(ttlData)) !== null) {
    const prefix = prefixMatch[1]; // e.g., "iriscategory" or "" for default
    const uri = prefixMatch[2]; // e.g., "https://codelist.commonapproach.org/IRISImpactCategory#"
    prefixMap[prefix] = uri;
  }

  console.log(`\n=== Parsing Turtle: ${sourceUrl} ===`);
  console.log(`Base URI: ${baseUri}`);
  console.log("Prefix map:", prefixMap);
  console.log(`Data size: ${(ttlData.length / 1024).toFixed(2)} KB`);

  // 🔥 CRITICAL FIX: Check if this is IRIS file (uses full iris.thegiin.org URLs)
  const isIRISFile = ttlData.includes("iris.thegiin.org");

  if (isIRISFile) {
    console.log("🎯 Detected IRIS file - using full URL parsing");
  }

  const lines = ttlData.split("\n");
  let currentEntry: CodeList | null = null;
  let currentBlock = "";
  let entryCount = 0;
  let skippedCount = 0;

  for (let i = 0; i < lines.length; i++) {
    const originalLine = lines[i];
    const trimmedLine = originalLine.trim();

    // Skip empty lines, comments, and prefix declarations
    if (
      !trimmedLine ||
      trimmedLine.startsWith("#") ||
      trimmedLine.startsWith("@prefix") ||
      trimmedLine.startsWith("@base")
    ) {
      continue;
    }

    // 🔥 CRITICAL: Try FULL URL format FIRST (for IRIS)
    const fullUrlMatch = trimmedLine.match(/<(https?:\/\/[^>]+)>\s*$/);

    // Try PREFIX notation format (EDG, Corporate)
    const localPrefixMatch = !fullUrlMatch ? trimmedLine.match(/^:([a-zA-Z0-9_-]+)(?:\s|$)/) : null;

    // Try EXTERNAL PREFIX notation format (iriscategory:Agriculture)
    // Matches prefix:Identifier where prefix is alphanumeric
    let externalPrefixMatch =
      !fullUrlMatch && !localPrefixMatch
        ? trimmedLine.match(/^([a-zA-Z0-9_-]+:[a-zA-Z0-9_-]+)(?:\s|$)/)
        : null;

    // Filter out known property prefixes to avoid treating properties as new entries
    if (externalPrefixMatch) {
      const prefix = externalPrefixMatch[1].split(":")[0];
      const IGNORED_PREFIXES = [
        "skos",
        "org",
        "dcterms",
        "cids",
        "rdf",
        "rdfs",
        "owl",
        "voaf",
        "xsd",
        "vann",
        "prov",
      ];
      if (IGNORED_PREFIXES.includes(prefix)) {
        externalPrefixMatch = null;
      }
    }

    if (fullUrlMatch || localPrefixMatch || externalPrefixMatch) {
      // Save previous entry if it exists
      if (currentEntry && currentEntry.hasName) {
        if (!isMetadataEntry(currentEntry.hasIdentifier, currentEntry.hasName)) {
          codeList.push(currentEntry);
          entryCount++;
          console.log(
            `  ✅ Entry ${entryCount}: ${currentEntry.hasIdentifier} - "${currentEntry.hasName}"`
          );
        } else {
          skippedCount++;
          console.log(`  ⏭️  Skipped metadata ${skippedCount}: ${currentEntry.hasIdentifier}`);
        }
      }

      // Start new entry
      if (fullUrlMatch) {
        // Full URL format (IRIS)
        const fullUrl = fullUrlMatch[1];
        const identifier = fullUrl.split("/").filter(Boolean).pop() || fullUrl;

        currentEntry = {
          "@id": fullUrl,
          hasIdentifier: identifier,
          hasName: "",
        };
        console.log(`  🔍 Found full URL: ${identifier}`);
      } else if (localPrefixMatch) {
        // Local Prefix notation format (:Identifier)
        const identifier = localPrefixMatch[1];
        currentEntry = {
          "@id": baseUri + identifier,
          hasIdentifier: identifier,
          hasName: "",
        };
      } else if (externalPrefixMatch) {
        // External Prefix notation format (prefix:Identifier)
        const rawId = externalPrefixMatch[1]; // e.g. iriscategory:Agriculture
        const [prefix, localName] = rawId.split(":");
        const identifier = localName || rawId;

        // 🔥 Expand the prefix to full URI using the prefix map
        let fullUri = rawId; // default fallback
        if (prefixMap[prefix]) {
          fullUri = prefixMap[prefix] + localName;
          console.log(`  🔗 Expanded ${rawId} → ${fullUri}`);
        }

        currentEntry = {
          "@id": fullUri,
          hasIdentifier: identifier,
          hasName: "",
        };
      }

      currentBlock = trimmedLine;
    } else if (currentEntry) {
      // Continue building current entry's block
      currentBlock += " " + trimmedLine;
    }

    // Extract properties from accumulated block
    if (currentEntry && currentBlock) {
      // Override hasIdentifier if explicitly defined
      const identifierMatch = currentBlock.match(
        /(?:cids:hasIdentifier|org:hasIdentifier|skos:notation)\s+"([^"]+)"/
      );
      if (identifierMatch) {
        currentEntry.hasIdentifier = identifierMatch[1];
      }

      // Extract @type
      if (!currentEntry["@type"]) {
        const typeMatch = currentBlock.match(/(?:\sa\s|rdf:type\s)([^;.]+)/);
        if (typeMatch) {
          let typeVal = typeMatch[1].trim();
          // Handle multiple types (comma separated)
          // Clean up whitespace and newlines, ensure comma separation is clean
          typeVal = typeVal.replace(/\s+/g, " ");
          typeVal = typeVal
            .split(",")
            .map((t) => t.trim())
            .join(", ");

          currentEntry["@type"] = typeVal;
        }
      }

      // Extract hasName from multiple possible predicates
      if (!currentEntry.hasName) {
        const nameMatch = currentBlock.match(
          /(?:cids:hasName|org:hasName|rdfs:label|skos:prefLabel)\s+"([^"]+)"(?:@[a-z-]+)?/
        );
        if (nameMatch) {
          currentEntry.hasName = nameMatch[1];
        }
      }

      // Extract description from multiple possible predicates
      if (!currentEntry.hasDescription) {
        const descMatch = currentBlock.match(
          /(?:cids:hasDescription|cids:hasDefinition|cids:hasCharacteristic|skos:definition)\s+"([^"]+)"(?:@[a-z-]+)?/
        );
        if (descMatch) {
          currentEntry.hasDescription = descMatch[1];
        }
      }
    }
  }

  // Don't forget the last entry
  if (currentEntry && currentEntry.hasName) {
    if (!isMetadataEntry(currentEntry.hasIdentifier, currentEntry.hasName)) {
      codeList.push(currentEntry);
      entryCount++;
      console.log(
        `  ✅ Entry ${entryCount}: ${currentEntry.hasIdentifier} - "${currentEntry.hasName}"`
      );
    } else {
      skippedCount++;
      console.log(`  ⏭️  Skipped metadata: ${currentEntry.hasIdentifier}`);
    }
  }

  console.log(`📊 Total parsed: ${codeList.length} entries`);
  console.log("=== End Parsing ===\n");

  return codeList;
}

// ============================================================================
// FETCH AND PARSE
// ============================================================================

async function fetchAndParseCodeList(url: string): Promise<CodeList[]> {
  try {
    const cachedData = getCachedData(url);
    if (cachedData) {
      return cachedData;
    }

    console.log(`🌐 Fetching: ${url}`);
    let data: string;
    let codeList: CodeList[] = [];

    try {
      const response = await fetch(url);
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }
      data = await response.text();
      console.log(`✅ Fetch successful (${(data.length / 1024).toFixed(2)} KB)`);
    } catch (primaryError) {
      console.warn(`⚠️ Primary fetch failed: ${(primaryError as Error).message}`);

      const fallbackUrl = GITHUB_FALLBACK_URLS[url];
      if (!fallbackUrl) {
        throw new Error(`No fallback URL available for ${url}`);
      }

      console.log(`🔄 Trying GitHub fallback: ${fallbackUrl}`);
      const fallbackResponse = await fetch(fallbackUrl);
      if (!fallbackResponse.ok) {
        console.error(`❌ Fallback also failed (HTTP ${fallbackResponse.status}) for ${url}`);
        return []; // Return empty array instead of crashing
      }
      data = await fallbackResponse.text();
      console.log("✅ Fallback successful");
    }

    // Parse based on file extension
    if (url.endsWith(".ttl")) {
      codeList = parseTurtleToCodeList(data, url);
    } else if (url.endsWith(".owl")) {
      codeList = parseXmlToCodeList(data);
    } else {
      throw new Error(`Unsupported file format for ${url}`);
    }

    if (codeList.length > 0) {
      setCachedData(url, codeList);
    } else {
      console.warn(`⚠️ Warning: Parsed 0 entries from ${url}`);
    }

    return codeList;
  } catch (error) {
    console.error(`❌ Failed to fetch and parse ${url}:`, error);
    throw error;
  }
}

// ============================================================================
// PUBLIC API
// ============================================================================

export async function getAllSectors(): Promise<CodeList[]> {
  try {
    console.log("\n🌍 === FETCHING ALL SECTORS === 🌍");

    // Only pull in the newly revised ESDCSector list
    const esdc = await fetchAndParseCodeList(CODELIST_URLS.ESDCSector);

    console.log(`\n✨ Total Sectors: ${esdc.length}\n`);

    return esdc;
  } catch (error) {
    console.error("❌ Error in getAllSectors():", error);
    return [];
  }
}

export async function getAllPopulationServed(): Promise<CodeList[]> {
  try {
    return await fetchAndParseCodeList(CODELIST_URLS.PopulationServed);
  } catch (error) {
    console.error("❌ Error fetching PopulationServed:", error);
    return [];
  }
}

export async function getAllProvinceTerritory(): Promise<CodeList[]> {
  try {
    return await fetchAndParseCodeList(CODELIST_URLS.ProvinceTerritory);
  } catch (error) {
    console.error("❌ Error fetching ProvinceTerritory:", error);
    return [];
  }
}

export async function getAllFundingState(): Promise<CodeList[]> {
  try {
    return await fetchAndParseCodeList(CODELIST_URLS.FundingState);
  } catch (error) {
    console.error("❌ Error fetching FundingState:", error);
    return [];
  }
}

export async function getAllSDGImpacts(): Promise<CodeList[]> {
  try {
    return await fetchAndParseCodeList(CODELIST_URLS.SDGImpacts);
  } catch (error) {
    console.error("❌ Error fetching SDGImpacts:", error);
    return [];
  }
}

export async function getAllOrganizationType(): Promise<CodeList[]> {
  try {
    return await fetchAndParseCodeList(CODELIST_URLS.OrganizationType);
  } catch (error) {
    console.error("❌ Error fetching OrganizationType:", error);
    return [];
  }
}

export async function getAllLocalities(): Promise<CodeList[]> {
  try {
    return await fetchAndParseCodeList(CODELIST_URLS.Locality);
  } catch (error) {
    console.error("❌ Error fetching Locality:", error);
    return [];
  }
}

export async function getAllCorporateRegistrars(): Promise<CodeList[]> {
  try {
    return await fetchAndParseCodeList(CODELIST_URLS.CorporateRegistrar);
  } catch (error) {
    console.error("❌ Error fetching CorporateRegistrar:", error);
    return [];
  }
}

export async function getAllIRISImpactCategories(): Promise<CodeList[]> {
  try {
    return await fetchAndParseCodeList(CODELIST_URLS.IRISImpactCategory);
  } catch (error) {
    console.error("❌ Error fetching IRIS Impact Categories:", error);
    return [];
  }
}

export async function getCodeListByTableName(tableName: string): Promise<CodeList[]> {
  const urlMap: Record<string, string> = {
    Sector: CODELIST_URLS.ESDCSector,
    PopulationServed: CODELIST_URLS.PopulationServed,
    Locality: CODELIST_URLS.Locality,
    ProvinceTerritory: CODELIST_URLS.ProvinceTerritory,
    FundingState: CODELIST_URLS.FundingState,
    SDGImpacts: CODELIST_URLS.SDGImpacts,
    OrganizationType: CODELIST_URLS.OrganizationType,
    CorporateRegistrar: CODELIST_URLS.CorporateRegistrar,
    IRISImpactCategory: CODELIST_URLS.IRISImpactCategory,
  };

  const url = urlMap[tableName];
  if (!url) {
    throw new Error(`No codelist URL found for table: ${tableName}`);
  }

  return fetchAndParseCodeList(url);
}
