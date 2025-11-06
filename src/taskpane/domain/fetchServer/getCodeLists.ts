import { XMLParser } from "fast-xml-parser";

// ============================================================================
// TYPE DEFINITIONS
// ============================================================================

export interface CodeList {
	"@id": string;
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

/** Primary codelist URLs from commonapproach.org */
const CODELIST_URLS = {
	ICNPOSector: "https://codelist.commonapproach.org/ICNPOsector/ICNPOsector.owl",
	StatsCanSector: "https://codelist.commonapproach.org/StatsCanSector/StatsCanSector.owl",
	PopulationServed: "https://codelist.commonapproach.org/PopulationServed/PopulationServed.owl",
	ProvinceTerritory: "https://codelist.commonapproach.org/ProvinceTerritory/ProvinceTerritory.owl",
	OrganizationType: "https://codelist.commonapproach.org/OrgTypeGOC/OrgTypeGOC.owl",
	Locality: "https://codelist.commonapproach.org/Locality/LocalityStatsCan.owl",
	CorporateRegistrar: "https://codelist.commonapproach.org/CanadianCorporateRegistries/CanadianCorporateRegistries.ttl",
	IRISImpactCategory: "https://codelist.commonapproach.org/IRISImpactThemes/IRISImpactCategories.ttl",
	EquityDeservingGroup: "https://codelist.commonapproach.org/EquityDeservingGroupsESDC/EquityDeservingGroupsESDC.ttl",
} as const;

/** GitHub fallback URLs for redundancy */
const GITHUB_FALLBACK_URLS: Record<string, string> = {
	[CODELIST_URLS.ICNPOSector]: 
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/ICNPOsector/ICNPOsector.owl",
	[CODELIST_URLS.StatsCanSector]: 
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/StatsCanSector/StatsCanSector.owl",
	[CODELIST_URLS.PopulationServed]: 
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/PopulationServed/PopulationServed.owl",
	[CODELIST_URLS.ProvinceTerritory]: 
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/ProvinceTerritory/ProvinceTerritory.owl",
	[CODELIST_URLS.OrganizationType]: 
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/OrgTypeGOC/OrgTypeGOC.owl",
	[CODELIST_URLS.Locality]: 
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/Locality/LocalityStatsCan.owl",
	[CODELIST_URLS.CorporateRegistrar]: 
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/CanadianCorporateRegistries/CanadianCorporateRegistries.ttl",
	[CODELIST_URLS.IRISImpactCategory]: 
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/IRISImpactCategories/IRISImpactCategories.ttl",
	[CODELIST_URLS.EquityDeservingGroup]: 
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/EquityDeservingGroupsESDC/EquityDeservingGroupsESDC.ttl",
};

/** Metadata identifiers to skip during parsing */
const METADATA_IDENTIFIERS = new Set([
	'dataset',
	'IRISImpactCategories',
	'CanadianCorporateRegistries',
	'EquityDeservingGroupsESDC',
	'ICNPOsector',
	'StatsCanSector',
	'PopulationServed',
	'ProvinceTerritory',
	'OrgTypeGOC',
	'LocalityStatsCan',
]);

/** Keywords that indicate a metadata entry */
const METADATA_KEYWORDS = ['Codelist', 'Code List', 'Categories', 'Registries', 'Dataset'];

// ============================================================================
// CACHE MANAGEMENT
// ============================================================================

const inMemoryCache: Record<string, CodeList[]> = {};

/**
 * Retrieves cached data from memory or localStorage
 */
function getCachedData(url: string): CodeList[] | null {
	// Check in-memory cache first (fastest)
	if (inMemoryCache[url]?.length > 0) {
		console.log(`✅ Cache hit (memory): ${url}`);
		return inMemoryCache[url];
	}

	// Check localStorage
	const cachedData = localStorage.getItem(url);
	if (!cachedData) {
		return null;
	}

	try {
		const parsedData = JSON.parse(cachedData) as CacheItem;

		// Validate cache structure
		if (!parsedData.data || !parsedData.timestamp || !parsedData.expiresIn) {
			console.warn(`⚠️ Invalid cache structure for ${url}, removing`);
			localStorage.removeItem(url);
			return null;
		}

		// Check expiration
		const isExpired = Date.now() - parsedData.timestamp > parsedData.expiresIn;
		if (isExpired) {
			console.log(`⏰ Cache expired for ${url}, removing`);
			localStorage.removeItem(url);
			return null;
		}

		// Cache is valid - store in memory for faster access
		console.log(`✅ Cache hit (localStorage): ${url}`);
		inMemoryCache[url] = parsedData.data;
		return parsedData.data;
	} catch (error) {
		console.error(`❌ Error parsing cached data for ${url}:`, error);
		localStorage.removeItem(url);
		return null;
	}
}

/**
 * Stores data in both memory and localStorage cache
 */
function setCachedData(url: string, data: CodeList[]): void {
	if (!data || data.length === 0) {
		return;
	}

	// Store in memory cache
	inMemoryCache[url] = data;

	// Store in localStorage with expiration
	const cacheItem: CacheItem = {
		data,
		timestamp: Date.now(),
		expiresIn: CACHE_EXPIRATION_TIME,
	};

	try {
		localStorage.setItem(url, JSON.stringify(cacheItem));
		console.log(`💾 Cached ${data.length} entries for ${url}`);
	} catch (error) {
		console.warn(`⚠️ Failed to cache in localStorage (${url}):`, error);
		// Continue execution - caching failure shouldn't break functionality
	}
}

/**
 * Clears cache for a specific table or all tables
 */
export function clearCodeListCache(tableName?: string): void {
	if (tableName) {
		// Map table name to URL
		const urlMap: Record<string, string> = {
			Sector: CODELIST_URLS.ICNPOSector, // Sector uses multiple URLs, clear the first one
			PopulationServed: CODELIST_URLS.PopulationServed,
			Locality: CODELIST_URLS.Locality,
			ProvinceTerritory: CODELIST_URLS.ProvinceTerritory,
			OrganizationType: CODELIST_URLS.OrganizationType,
			CorporateRegistrar: CODELIST_URLS.CorporateRegistrar,
			IRISImpactCategory: CODELIST_URLS.IRISImpactCategory,
			EquityDeservingGroup: CODELIST_URLS.EquityDeservingGroup,
		};

		const url = urlMap[tableName];
		if (url) {
			delete inMemoryCache[url];
			localStorage.removeItem(url);
			console.log(`🗑️ Cleared cache for ${tableName}`);
		}
	} else {
		// Clear all caches
		Object.keys(inMemoryCache).forEach(key => delete inMemoryCache[key]);
		Object.keys(localStorage).forEach(key => {
			if (key.includes('codelist.commonapproach.org')) {
				localStorage.removeItem(key);
			}
		});
		console.log("🗑️ Cleared all codelist caches");
	}
}

// ============================================================================
// XML PARSER (for .owl files)
// ============================================================================

/**
 * Parses OWL/RDF XML format to CodeList array
 * Used for: ICNPO, StatsCan, PopulationServed, ProvinceTerritory, OrganizationType, Locality
 */
function parseXmlToCodeList(xmlData: string): CodeList[] {
	const parser = new XMLParser({ ignoreAttributes: false });
	const jsonData = parser.parse(xmlData);

	const codeList: CodeList[] = [];
	const descriptions = jsonData["rdf:RDF"]?.["rdf:Description"] || [];
	let baseIdUrl = "";

	// Ensure descriptions is an array
	const descArray = Array.isArray(descriptions) ? descriptions : [descriptions];

	for (const desc of descArray) {
		// Extract base URL from first entry
		if (desc["vann:preferredNamespacePrefix"]) {
			baseIdUrl = desc["@_rdf:about"]?.replace("#dataset", "") || "";
			continue;
		}

		// Skip entries without required fields
		if (!desc["cids:hasIdentifier"] && !desc["cids:hasName"]) {
			continue;
		}

		// Build CodeList entry
		const entry: CodeList = {
			"@id": desc["@_rdf:about"]?.includes(baseIdUrl)
				? desc["@_rdf:about"]
				: baseIdUrl + desc["@_rdf:about"],
			hasIdentifier: desc["cids:hasIdentifier"]?.toString() || "",
			hasName: desc["cids:hasName"]?.["#text"]?.toString() || "",
		};

		// Extract description from multiple possible predicates
		if (desc["cids:hasDescription"]?.["#text"]) {
			entry.hasDescription = desc["cids:hasDescription"]["#text"].toString();
		} else if (desc["cids:hasDefinition"]?.["#text"]) {
			entry.hasDescription = desc["cids:hasDefinition"]["#text"].toString();
		} else if (desc["cids:hasCharacteristic"]?.["#text"]) {
			entry.hasDescription = desc["cids:hasCharacteristic"]["#text"].toString();
		}

		codeList.push(entry);
	}

	return codeList;
}

// ============================================================================
// TURTLE PARSER (for .ttl files)
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
	return METADATA_KEYWORDS.some(keyword => name.includes(keyword));
}

/**
 * Enhanced Turtle parser supporting multiple formats:
 * 1. Full URLs: <https://iris.thegiin.org/theme/5.3/Agriculture/> (IRIS)
 * 2. Prefix notation: :identifier (CorporateRegistrar, EDG)
 * 
 * Handles multiple description predicates:
 * - cids:hasDescription
 * - cids:hasDefinition
 * - cids:hasCharacteristic (EDG)
 * - skos:definition (IRIS)
 */
function parseTurtleToCodeList(ttlData: string, sourceUrl: string): CodeList[] {
	const codeList: CodeList[] = [];
	
	// Extract base URI from @prefix declaration
	let baseUri = "https://codelist.commonapproach.org/codeLists/";
	const baseUriMatch = ttlData.match(/@prefix\s*:\s*<([^>]+)>/m);
	if (baseUriMatch) {
		baseUri = baseUriMatch[1];
	}

	console.log(`\n=== Parsing Turtle: ${sourceUrl} ===`);
	console.log(`Base URI: ${baseUri}`);
	console.log(`Data size: ${(ttlData.length / 1024).toFixed(2)} KB`);

	// 🔥 CRITICAL FIX: Check if this is IRIS file (uses full iris.thegiin.org URLs)
	const isIRISFile = ttlData.includes('iris.thegiin.org');
	
	if (isIRISFile) {
		console.log(`🎯 Detected IRIS file - using full URL parsing`);
	}

	const lines = ttlData.split('\n');
	let currentEntry: CodeList | null = null;
	let currentBlock = '';
	let entryCount = 0;
	let skippedCount = 0;

	for (let i = 0; i < lines.length; i++) {
		const line = lines[i].trim();

		// Skip empty lines, comments, and prefix declarations
		if (!line || line.startsWith('#') || line.startsWith('@prefix') || line.startsWith('@base')) {
			continue;
		}

		// 🔥 CRITICAL: Try FULL URL format FIRST (for IRIS)
		const fullUrlMatch = line.match(/<(https?:\/\/[^>]+)>\s*$/);
		
		// Try PREFIX notation format (EDG, Corporate)
		const prefixMatch = !fullUrlMatch ? line.match(/^:([a-zA-Z0-9_-]+)(?:\s|$)/) : null;

		if (fullUrlMatch || prefixMatch) {
			// Save previous entry if it exists
			if (currentEntry && currentEntry.hasName) {
				if (!isMetadataEntry(currentEntry.hasIdentifier, currentEntry.hasName)) {
					codeList.push(currentEntry);
					entryCount++;
					console.log(`  ✅ Entry ${entryCount}: ${currentEntry.hasIdentifier} - "${currentEntry.hasName}"`);
				} else {
					skippedCount++;
					console.log(`  ⏭️  Skipped metadata: ${currentEntry.hasIdentifier}`);
				}
			}

			// Start new entry
			if (fullUrlMatch) {
				// Full URL format (IRIS)
				const fullUrl = fullUrlMatch[1];
				const identifier = fullUrl.split('/').filter(Boolean).pop() || fullUrl;
				
				currentEntry = {
					"@id": fullUrl,
					hasIdentifier: identifier,
					hasName: "",
				};
				console.log(`  🔍 Found full URL: ${identifier}`);
			} else if (prefixMatch) {
				// Prefix notation format
				const identifier = prefixMatch[1];
				currentEntry = {
					"@id": baseUri + identifier,
					hasIdentifier: identifier,
					hasName: "",
				};
			}

			currentBlock = line;
		} else if (currentEntry) {
			// Continue building current entry's block
			currentBlock += ' ' + line;
		}

		// Extract properties from accumulated block
		if (currentEntry && currentBlock) {
			// Override hasIdentifier if explicitly defined
			const identifierMatch = currentBlock.match(/cids:hasIdentifier\s+"([^"]+)"/);
			if (identifierMatch) {
				currentEntry.hasIdentifier = identifierMatch[1];
			}

			// Extract hasName from multiple possible predicates
			if (!currentEntry.hasName) {
				const nameMatch = currentBlock.match(/(?:cids:hasName|rdfs:label)\s+"([^"]+)"(?:@[a-z-]+)?/);
				if (nameMatch) {
					currentEntry.hasName = nameMatch[1];
				}
			}

			// Extract description from multiple possible predicates
			if (!currentEntry.hasDescription) {
				const descMatch = currentBlock.match(/(?:cids:hasDescription|cids:hasDefinition|cids:hasCharacteristic|skos:definition)\s+"([^"]+)"(?:@[a-z-]+)?/);
				if (descMatch) {
					currentEntry.hasDescription = descMatch[1];
				}
			}
		}
	}

	// Don't forget the last entry!
	if (currentEntry && currentEntry.hasName) {
		if (!isMetadataEntry(currentEntry.hasIdentifier, currentEntry.hasName)) {
			codeList.push(currentEntry);
			entryCount++;
			console.log(`  ✅ Entry ${entryCount}: ${currentEntry.hasIdentifier} - "${currentEntry.hasName}"`);
		} else {
			skippedCount++;
			console.log(`  ⏭️  Skipped metadata: ${currentEntry.hasIdentifier}`);
		}
	}

	console.log(`📊 Parsed ${entryCount} entries (skipped ${skippedCount} metadata entries)`);
	console.log(`=== End Parsing ===\n`);

	return codeList;
}

// ============================================================================
// FETCH AND PARSE
// ============================================================================

/**
 * Fetches and parses a codelist from URL with automatic fallback
 */
async function fetchAndParseCodeList(url: string): Promise<CodeList[]> {
	try {
		// Check cache first
		const cachedData = getCachedData(url);
		if (cachedData) {
			return cachedData;
		}

		console.log(`🌐 Fetching: ${url}`);
		let data: string;
		let codeList: CodeList[] = [];
		let fetchError: Error | null = null;

		// Try primary URL
		try {
			const response = await fetch(url);
			if (!response.ok) {
				throw new Error(`HTTP ${response.status}: ${response.statusText}`);
			}
			data = await response.text();
			console.log(`✅ Primary fetch successful (${(data.length / 1024).toFixed(2)} KB)`);
		} catch (primaryError) {
			fetchError = primaryError as Error;
			console.warn(`⚠️ Primary fetch failed: ${fetchError.message}`);

			// Try GitHub fallback
			const fallbackUrl = GITHUB_FALLBACK_URLS[url];
			if (!fallbackUrl) {
				throw new Error(`No fallback URL available for ${url}`);
			}

			console.log(`🔄 Trying GitHub fallback: ${fallbackUrl}`);
			const fallbackResponse = await fetch(fallbackUrl);
			if (!fallbackResponse.ok) {
				throw new Error(`Fallback HTTP ${fallbackResponse.status}: ${fallbackResponse.statusText}`);
			}
			data = await fallbackResponse.text();
			console.log(`✅ Fallback fetch successful (${(data.length / 1024).toFixed(2)} KB)`);
		}

		// Parse based on file extension
		if (url.endsWith('.ttl')) {
			codeList = parseTurtleToCodeList(data, url);
		} else if (url.endsWith('.owl')) {
			codeList = parseXmlToCodeList(data);
		} else {
			throw new Error(`Unsupported file format for ${url}`);
		}

		// Cache successful results
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
// PUBLIC API - Individual Codelist Fetchers
// ============================================================================

export async function getAllSectors(): Promise<CodeList[]> {
	try {
		console.log("\n🌍 === FETCHING ALL SECTORS === 🌍");
		
		const [icnpo, statsCan, iris] = await Promise.all([
			fetchAndParseCodeList(CODELIST_URLS.ICNPOSector),
			fetchAndParseCodeList(CODELIST_URLS.StatsCanSector),
			fetchAndParseCodeList(CODELIST_URLS.IRISImpactCategory),
		]);

		const combined = [...icnpo, ...statsCan, ...iris];
		console.log(`\n✨ Total Sectors: ${combined.length} (ICNPO: ${icnpo.length}, StatsCan: ${statsCan.length}, IRIS: ${iris.length})\n`);
		
		return combined;
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

export async function getAllEquityDeservingGroups(): Promise<CodeList[]> {
	try {
		return await fetchAndParseCodeList(CODELIST_URLS.EquityDeservingGroup);
	} catch (error) {
		console.error("❌ Error fetching Equity Deserving Groups:", error);
		return [];
	}
}

// ============================================================================
// PUBLIC API - Generic Fetcher by Table Name
// ============================================================================

/**
 * Fetches codelist data by table name
 * Special handling: "Sector" combines ICNPO, StatsCan, and IRIS
 */
export async function getCodeListByTableName(tableName: string): Promise<CodeList[]> {
	// Special case: Sector combines three codelists
	if (tableName === "Sector") {
		return getAllSectors();
	}

	// Map table names to URLs
	const urlMap: Record<string, string> = {
		PopulationServed: CODELIST_URLS.PopulationServed,
		Locality: CODELIST_URLS.Locality,
		ProvinceTerritory: CODELIST_URLS.ProvinceTerritory,
		OrganizationType: CODELIST_URLS.OrganizationType,
		CorporateRegistrar: CODELIST_URLS.CorporateRegistrar,
		IRISImpactCategory: CODELIST_URLS.IRISImpactCategory,
		EquityDeservingGroup: CODELIST_URLS.EquityDeservingGroup,
	};

	const url = urlMap[tableName];
	if (!url) {
		throw new Error(`No codelist URL found for table: ${tableName}`);
	}

	return fetchAndParseCodeList(url);
}