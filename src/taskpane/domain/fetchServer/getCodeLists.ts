import { XMLParser } from "fast-xml-parser";

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

const inMemoryCache: { [key: string]: CodeList[] } = {};
const CACHE_EXPIRATION_TIME = 24 * 60 * 60 * 1000; // 24 hours

// GitHub fallback URLs mapping
const GITHUB_FALLBACK_URLS: { [key: string]: string } = {
	"https://codelist.commonapproach.org/ICNPOsector/ICNPOsector.owl":
		"https://raw.githubusercontent.com/commonapproach.org/CodeLists/main/ICNPOsector/ICNPOsector.owl",
	"https://codelist.commonapproach.org/StatsCanSector/StatsCanSector.owl":
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/StatsCanSector/StatsCanSector.owl",
	"https://codelist.commonapproach.org/PopulationServed/PopulationServed.owl":
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/PopulationServed/PopulationServed.owl",
	"https://codelist.commonapproach.org/ProvinceTerritory/ProvinceTerritory.owl":
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/ProvinceTerritory/ProvinceTerritory.owl",
	"https://codelist.commonapproach.org/OrgTypeGOC/OrgTypeGOC.owl":
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/OrgTypeGOC/OrgTypeGOC.owl",
	"https://codelist.commonapproach.org/Locality/LocalityStatsCan.owl":
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/Locality/LocalityStatsCan.owl",
	"https://codelist.commonapproach.org/CanadianCorporateRegistries/CanadianCorporateRegistries.ttl":
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/CanadianCorporateRegistries/CanadianCorporateRegistries.ttl",	
	"https://codelist.commonapproach.org/IRISImpactThemes/IRISImpactCategories.ttl":
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/IRISImpactCategories/IRISImpactCategories.ttl",
	"https://codelist.commonapproach.org/EquityDeservingGroupsESDC/EquityDeservingGroupsESDC.ttl":
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/EquityDeservingGroupsESDC/EquityDeservingGroupsESDC.ttl",
};

function parseXmlToCodeList(data: string): CodeList[] {
	const options = {
		ignoreAttributes: false,
	};

	const parser = new XMLParser(options);
	const jsonData = parser.parse(data);

	const codeList: CodeList[] = [];
	const descriptions = jsonData["rdf:RDF"]["rdf:Description"] || [];
	let baseIdUrl = "";

	for (let desc of descriptions) {
		if (desc["vann:preferredNamespacePrefix"]) {
			baseIdUrl = desc["@_rdf:about"].replace("#dataset", "");
			continue;
		}

		if (!desc["cids:hasIdentifier"] && !desc["cids:hasName"]) {
			continue;
		}

		const sector: CodeList = {
			"@id": desc["@_rdf:about"].includes(baseIdUrl)
				? desc["@_rdf:about"]
				: baseIdUrl + desc["@_rdf:about"],
			hasIdentifier: desc["cids:hasIdentifier"] ? desc["cids:hasIdentifier"].toString() : "",
			hasName: desc["cids:hasName"]["#text"] ? desc["cids:hasName"]["#text"].toString() : "",
		};

		// Try multiple predicates for description field (mapping to hasDescription in output)
		if (desc["cids:hasDescription"]) {
			sector.hasDescription = desc["cids:hasDescription"]["#text"]
				? desc["cids:hasDescription"]["#text"].toString()
				: "";
		} else if (desc["cids:hasDefinition"]) {
			sector.hasDescription = desc["cids:hasDefinition"]["#text"]
				? desc["cids:hasDefinition"]["#text"].toString()
				: "";
		} else if (desc["cids:hasCharacteristic"]) {
			// EDG uses hasCharacteristic instead of hasDescription
			sector.hasDescription = desc["cids:hasCharacteristic"]["#text"]
				? desc["cids:hasCharacteristic"]["#text"].toString()
				: "";
		}

		codeList.push(sector);
	}

	return codeList;
}

/**
 * Enhanced Turtle parser that handles multiple description predicates
 * Maps hasDefinition, hasCharacteristic, and other predicates to hasDescription
 */
function parseTurtleToCodeList(ttlData: string): CodeList[] {
	const codeList: CodeList[] = [];
	
	// Extract base URI
	let baseUri = "https://codelist.commonapproach.org/codeLists/";
	const baseUriMatch = ttlData.match(/@prefix\s*:\s*<([^>]+)>/m);
	if (baseUriMatch) {
		baseUri = baseUriMatch[1];
	}

	console.log("=== PARSING TURTLE ===");
	console.log("Base URI:", baseUri);

	// Split into individual entries by looking for lines starting with :id
	const lines = ttlData.split('\n');
	let currentEntry: CodeList | null = null;
	let currentBlock = '';
	
	for (let i = 0; i < lines.length; i++) {
		const line = lines[i].trim();
		
		// Skip empty lines, comments, and prefix declarations
		if (!line || line.startsWith('#') || line.startsWith('@prefix') || line.startsWith('@base')) {
			continue;
		}
		
		// Match entry start: :id (handles with or without trailing space)
		const entryMatch = line.match(/^:([a-zA-Z0-9_-]+)(?:\s|$)/);
		
		if (entryMatch) {
			// Save previous entry if it exists and has a name
			if (currentEntry && currentEntry.hasName) {
				codeList.push(currentEntry);
			}
			
			const id = entryMatch[1];
			
			// Skip dataset definition and codelist metadata headers
			// These are entries that describe the codelist itself, not actual code values
			// Examples: :IRISImpactCategories, :CanadianCorporateRegistries, :EquityDeservingGroupsESDC
			// Characteristics: Usually the ID matches the last part of the base URI
			const baseUriLastPart = baseUri.split('/').filter(p => p).pop() || '';
			const isMetadataEntry = id === 'dataset' || id === baseUriLastPart;
			
			if (isMetadataEntry) {
				console.log(`Skipping metadata entry: ${id}`);
				currentEntry = null;
				currentBlock = '';
				continue;
			}
			
			// Start new entry
			currentEntry = {
				"@id": baseUri + id,
				hasIdentifier: id, // Use ID as fallback identifier
				hasName: "",
			};
			currentBlock = line;
		} else if (currentEntry) {
			// Continue building current entry's block
			currentBlock += ' ' + line;
		}
		
		// Extract properties from the accumulated block
		if (currentEntry && currentBlock) {
			// Extract hasIdentifier (override fallback if found)
			const identifierMatch = currentBlock.match(/cids:hasIdentifier\s+"([^"]+)"/);
			if (identifierMatch) {
				currentEntry.hasIdentifier = identifierMatch[1];
			}
			
			// Extract hasName - try multiple predicates
			const nameMatch = currentBlock.match(/(?:cids:hasName|rdfs:label)\s+"([^"]+)"(?:@[a-z-]+)?/);
			if (nameMatch) {
				currentEntry.hasName = nameMatch[1];
			}
			
			// ✅ CRITICAL FIX: Extract description from MULTIPLE predicates
			// Map hasDefinition, hasCharacteristic, skos:definition to hasDescription
			if (!currentEntry.hasDescription) {
				const descMatch = currentBlock.match(/(?:cids:hasDescription|cids:hasDefinition|cids:hasCharacteristic|skos:definition)\s+"([^"]+)"(?:@[a-z-]+)?/);
				if (descMatch) {
					currentEntry.hasDescription = descMatch[1];
				}
			}
		}
	}
	
	// Don't forget to add the last entry
	if (currentEntry && currentEntry.hasName) {
		codeList.push(currentEntry);
	}
	
	console.log(`Total entries parsed: ${codeList.length}`);
	console.log("======================");
	
	return codeList;
}

async function fetchAndParseCodeList(url: string): Promise<CodeList[]> {
	try {
		// Check if the data is already in the in-memory cache
		if (inMemoryCache[url] && inMemoryCache[url].length > 0) {
			console.log(`Using in-memory cache for ${url}`);
			return inMemoryCache[url];
		}

		// Check if the data is in localStorage
		const cachedData = localStorage.getItem(url);
		if (cachedData) {
			try {
				const parsedData = JSON.parse(cachedData);

				// Check if it's the cache format with expiration
				if (parsedData.data && parsedData.timestamp && parsedData.expiresIn) {
					const now = Date.now();
					const isExpired = now - parsedData.timestamp > parsedData.expiresIn;

					if (!isExpired) {
						// Cache is still valid
						console.log(`Using localStorage cache for ${url}`);
						inMemoryCache[url] = parsedData.data;
						return parsedData.data;
					} else {
						// Cache expired, remove it
						console.log(`Cache expired for ${url}, removing`);
						localStorage.removeItem(url);
					}
				} else if (Array.isArray(parsedData)) {
					// Old cache format - invalidate it
					localStorage.removeItem(url);
				}
			} catch (error) {
				// Invalid JSON or corrupted cache, remove it
				localStorage.removeItem(url);
			}
		}

		let data: string;
		let codeList: CodeList[] = [];

		// Try to fetch from primary URL first
		try {
			console.log(`Attempting to fetch from primary URL: ${url}`);
			const response = await fetch(url);

			if (!response.ok) {
				throw new Error(`Primary fetch failed with status: ${response.status}`);
			}

			data = await response.text();
			
			// Parse based on file type
			if (url.endsWith('.ttl')) {
				codeList = parseTurtleToCodeList(data);
			} else {
				codeList = parseXmlToCodeList(data);
			}

			console.log(`✅ Successfully fetched ${codeList.length} items from primary URL`);
		} catch (primaryError) {
			console.warn(`Primary fetch failed for ${url}:`, primaryError);

			// Try GitHub fallback if available
			const fallbackUrl = GITHUB_FALLBACK_URLS[url];
			if (fallbackUrl) {
				try {
					console.log(`Attempting fallback from GitHub: ${fallbackUrl}`);
					const fallbackResponse = await fetch(fallbackUrl);

					if (!fallbackResponse.ok) {
						throw new Error(`Fallback fetch failed with status: ${fallbackResponse.status}`);
					}

					data = await fallbackResponse.text();
					
					// Parse based on file type
					if (url.endsWith('.ttl')) {
						codeList = parseTurtleToCodeList(data);
					} else {
						codeList = parseXmlToCodeList(data);
					}

					console.log(`✅ Successfully fetched ${codeList.length} items from GitHub fallback`);
				} catch (fallbackError) {
					console.error(`Both primary and fallback fetch failed for ${url}`);
					throw new Error(
						`All fetch attempts failed. Primary: ${primaryError}, Fallback: ${fallbackError}`
					);
				}
			} else {
				// No fallback available, re-throw primary error
				throw primaryError;
			}
		}

		// Cache the successful result
		if (codeList.length > 0) {
			// Store in memory cache
			inMemoryCache[url] = codeList;

			// Store in localStorage with expiration
			const cacheItem: CacheItem = {
				data: codeList,
				timestamp: Date.now(),
				expiresIn: CACHE_EXPIRATION_TIME,
			};
			try {
				localStorage.setItem(url, JSON.stringify(cacheItem));
				console.log(`Cached ${codeList.length} items for ${url}`);
			} catch (storageError) {
				console.warn("Failed to cache in localStorage:", storageError);
			}
		}

		return codeList;
	} catch (error) {
		console.error(`Error in fetchAndParseCodeList for ${url}:`, error);
		throw error;
	}
}

export async function getCodeListByTableName(tableName: string): Promise<CodeList[]> {
	const codelistUrls: { [key: string]: string } = {
		Sector: "https://codelist.commonapproach.org/ICNPOsector/ICNPOsector.owl",
		PopulationServed: "https://codelist.commonapproach.org/PopulationServed/PopulationServed.owl",
		Locality: "https://codelist.commonapproach.org/Locality/LocalityStatsCan.owl",
		ProvinceTerritory: "https://codelist.commonapproach.org/ProvinceTerritory/ProvinceTerritory.owl",
		OrganizationType: "https://codelist.commonapproach.org/OrgTypeGOC/OrgTypeGOC.owl",
		CorporateRegistrar: "https://codelist.commonapproach.org/CanadianCorporateRegistries/CanadianCorporateRegistries.ttl",
		IRISImpactCategory: "https://codelist.commonapproach.org/IRISImpactThemes/IRISImpactCategories.ttl",
		EquityDeservingGroup: "https://codelist.commonapproach.org/EquityDeservingGroupsESDC/EquityDeservingGroupsESDC.ttl",
	};

	const url = codelistUrls[tableName];
	if (!url) {
		throw new Error(`No codelist URL found for table: ${tableName}`);
	}

	return fetchAndParseCodeList(url);
}

export function clearCodeListCache(tableName?: string): void {
	if (tableName) {
		const codelistUrls: { [key: string]: string } = {
			Sector: "https://codelist.commonapproach.org/ICNPOsector/ICNPOsector.owl",
			PopulationServed: "https://codelist.commonapproach.org/PopulationServed/PopulationServed.owl",
			Locality: "https://codelist.commonapproach.org/Locality/LocalityStatsCan.owl",
			ProvinceTerritory: "https://codelist.commonapproach.org/ProvinceTerritory/ProvinceTerritory.owl",
			OrganizationType: "https://codelist.commonapproach.org/OrgTypeGOC/OrgTypeGOC.owl",
			CorporateRegistrar: "https://codelist.commonapproach.org/CanadianCorporateRegistries/CanadianCorporateRegistries.ttl",
			IRISImpactCategory: "https://codelist.commonapproach.org/IRISImpactThemes/IRISImpactCategories.ttl",
			EquityDeservingGroup: "https://codelist.commonapproach.org/EquityDeservingGroupsESDC/EquityDeservingGroupsESDC.ttl",
		};

		const url = codelistUrls[tableName];
		if (url) {
			delete inMemoryCache[url];
			localStorage.removeItem(url);
			console.log(`Cleared cache for ${tableName}`);
		}
	} else {
		// Clear all caches
		Object.keys(inMemoryCache).forEach(key => delete inMemoryCache[key]);
		Object.keys(localStorage).forEach(key => {
			if (key.includes('codelist.commonapproach.org')) {
				localStorage.removeItem(key);
			}
		});
		console.log("Cleared all codelist caches");
	}
}