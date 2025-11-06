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
		"https://raw.githubusercontent.com/commonapproach/CodeLists/main/ICNPOsector/ICNPOsector.owl",
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

		if (desc["cids:hasDescription"]) {
			sector.hasDescription = desc["cids:hasDescription"]["#text"]
				? desc["cids:hasDescription"]["#text"].toString()
				: "";
		} else if (desc["cids:hasDefinition"]) {
			sector.hasDescription = desc["cids:hasDefinition"]["#text"]
				? desc["cids:hasDefinition"]["#text"].toString()
				: "";
		}

		codeList.push(sector);
	}

	return codeList;
}

function parseTurtleToCodeList(ttlData: string): CodeList[] {
	const codeList: CodeList[] = [];
	
	// Extract base URI from @prefix
	let baseUri = "https://codelist.commonapproach.org/codeLists/";
	const baseUriMatch = ttlData.match(/@prefix\s*:\s*<([^>]+)>/m);
	if (baseUriMatch) {
		baseUri = baseUriMatch[1];
	}
	
	console.log("=== PARSING TURTLE ===");
	console.log("Base URI:", baseUri);
	
	// Split into lines and parse line-by-line
	const lines = ttlData.split('\n');
	let currentEntry: CodeList | null = null;
	let currentBlock = '';
	let fullUrlCount = 0;
	let prefixCount = 0;
	let skipCurrentEntry = false; // Flag to skip ConceptScheme entries
	
	for (let i = 0; i < lines.length; i++) {
		const line = lines[i];
		
		// Check if this line starts a new full-URL entry (IRIS format)
		if (line.match(/^<https?:\/\//)) {
			// Save previous entry if it exists and we're not skipping it
			if (currentEntry && currentEntry.hasName && !skipCurrentEntry) {
				codeList.push(currentEntry);
				fullUrlCount++;
				console.log(`Parsed full URL: ${currentEntry.hasName}`);
			}
			
			// Start new entry
			const urlMatch = line.match(/<(https?:\/\/[^>]+)>/);
			if (urlMatch) {
				currentEntry = {
					"@id": urlMatch[1],
					hasIdentifier: "",
					hasName: ""
				};
				currentBlock = line;
				skipCurrentEntry = false; // Reset skip flag
			}
		}
		// Check if this line starts a new prefix entry (:id format)
		else if (line.match(/^:[a-zA-Z0-9_-]+(?:\s|$)/)) {
			// Save previous entry if it exists and we're not skipping it
			if (currentEntry && currentEntry.hasName && !skipCurrentEntry) {
				codeList.push(currentEntry);
				prefixCount++;
				console.log(`Parsed prefix: ${currentEntry.hasName}`);
			}
			
			// Start new entry
			const idMatch = line.match(/:([a-zA-Z0-9_-]+)/);
			if (idMatch && idMatch[1] !== 'dataset') {
				currentEntry = {
					"@id": baseUri + idMatch[1],
					hasIdentifier: idMatch[1],
					hasName: ""
				};
				currentBlock = line;
				skipCurrentEntry = false; // Reset skip flag
			} else {
				currentEntry = null;
				currentBlock = '';
				skipCurrentEntry = false;
			}
		}
		else if (currentEntry && !skipCurrentEntry) {
			// Continue building current entry
			currentBlock += '\n' + line;
			
			// CHECK: If we see ConceptScheme in the "a" declaration, mark this entry to be skipped
			if (currentBlock.match(/\s+a\s+[^;]*skos:ConceptScheme/)) {
				console.log(`Skipping ConceptScheme: ${currentEntry["@id"]}`);
				skipCurrentEntry = true;
				currentEntry = null;
				currentBlock = '';
				continue;
			}
			
			// Check if we hit the end of this entry (line ending with ".")
			if (line.trim() === '.' || line.trim().endsWith(' .')) {
				// Extract properties from the complete block
				
				// hasName or rdfs:label (remove @en language tags)
				const nameMatch = currentBlock.match(/(?:cids:hasName|rdfs:label)\s+"([^"]+)"(?:@\w+)?/);
				if (nameMatch) {
					currentEntry.hasName = nameMatch[1];
				}
				
				// hasIdentifier (override if present)
				const identMatch = currentBlock.match(/cids:hasIdentifier\s+"([^"]+)"/);
				if (identMatch) {
					currentEntry.hasIdentifier = identMatch[1];
				} else if (!currentEntry.hasIdentifier) {
					// Use last part of URL as fallback
					const parts = currentEntry["@id"].split('/');
					currentEntry.hasIdentifier = parts.filter(p => p).pop() || "";
				}
				
				// hasDescription - handle cids:hasDescription, cids:hasDefinition, AND skos:definition
				// Use a flexible regex that handles escaped quotes and special characters
				const descMatch = currentBlock.match(/(?:cids:hasDescription|cids:hasDefinition|skos:definition)\s+"((?:[^"\\]|\\.)*)"/);
				if (descMatch) {
					// Unescape escaped characters
					currentEntry.hasDescription = descMatch[1]
						.replace(/\\"/g, '"')
						.replace(/\\\\/g, '\\');
				}
			}
		}
	}
	
	// ✅ CRITICAL: Don't forget the last entry! (only if not skipping)
	if (currentEntry && currentEntry.hasName && !skipCurrentEntry) {
		codeList.push(currentEntry);
		if (currentEntry["@id"].startsWith('http')) {
			fullUrlCount++;
		} else {
			prefixCount++;
		}
		console.log(`Parsed LAST entry: ${currentEntry.hasName}`);
	}
	
	console.log(`Full URL entries: ${fullUrlCount}, Prefix entries: ${prefixCount}`);
	console.log(`Total entries parsed: ${codeList.length}`);
	console.log("======================");
	
	return codeList;
}

async function fetchAndParseCodeList(url: string): Promise<CodeList[]> {
	try {
		// Check if the data is already in the in-memory cache
		if (inMemoryCache[url] && inMemoryCache[url].length > 0) {
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
						inMemoryCache[url] = parsedData.data;
						return parsedData.data;
					} else {
						// Cache expired, remove it
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

			console.log(`Successfully fetched ${codeList.length} items from primary URL`);
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

					console.log(`Successfully fetched ${codeList.length} items from GitHub fallback`);
				} catch (fallbackError) {
					console.error(`Both primary and fallback fetch failed for ${url}:`, fallbackError);
					throw new Error(
						`All fetch attempts failed. Primary: ${(primaryError as Error).message}, Fallback: ${(fallbackError as Error).message}`
					);
				}
			} else {
				console.error(`No fallback URL available for ${url}`);
				throw primaryError;
			}
		}

		// Store in in-memory cache
		inMemoryCache[url] = codeList;

		// Save to localStorage with expiration if not too large (< 200KB)
		if (codeList.length > 0) {
			const cacheItem: CacheItem = {
				data: codeList,
				timestamp: Date.now(),
				expiresIn: CACHE_EXPIRATION_TIME,
			};

			const serializedCache = JSON.stringify(cacheItem);
			if (serializedCache.length < 200000) {
				localStorage.setItem(url, serializedCache);
			}
		}

		return codeList;
	} catch (error) {
		console.error(`Error fetching or parsing ${url}:`, error);

		// Last resort: check for any cached data (even if expired)
		const lastResortCache = localStorage.getItem(url);
		if (lastResortCache) {
			try {
				const parsedData = JSON.parse(lastResortCache);
				if (parsedData.data && Array.isArray(parsedData.data)) {
					console.warn(`Using expired cached data for ${url} as fallback`);
					return parsedData.data;
				}
			} catch (cacheError) {
				console.error(`Failed to parse last resort cache for ${url}`);
			}
		}

		return [];
	}
}

export async function getCodeListByTableName(tableName: string): Promise<CodeList[]> {
	let codeList: CodeList[] = [];
	switch (tableName) {
		case "Sector":
			codeList = await getAllSectors();
			break;
		case "PopulationServed":
			codeList = await getAllPopulationServed();
			break;
		case "ProvinceTerritory":
			codeList = await getAllProvinceTerritory();
			break;
		case "OrganizationType":
			codeList = await getAllOrganizationType();
			break;
		case "Locality":
			codeList = await getAllLocalities();
			break;
		case "CorporateRegistrar":
			codeList = await getAllCorporateRegistrars();
			break;
		case "EquityDeservingGroup":
			codeList = await getAllEquityDeservingGroups();
			break;
		default:
			throw new Error(`Table ${tableName} not found`);
	}

	return codeList;
}

export async function getAllSectors(): Promise<CodeList[]> {
	try {
		const icnpoSectors = await fetchAndParseCodeList(
			"https://codelist.commonapproach.org/ICNPOsector/ICNPOsector.owl"
		);
		const statsCanSectors = await fetchAndParseCodeList(
			"https://codelist.commonapproach.org/StatsCanSector/StatsCanSector.owl"
		);
		const irisSectors = await fetchAndParseCodeList(
			"https://codelist.commonapproach.org/IRISImpactThemes/IRISImpactCategories.ttl" 
		);

		return [...icnpoSectors, ...statsCanSectors, ...irisSectors];
	} catch (error) {
		console.error("Error fetching sectors code list:", error);
		return [];
	}
}

export async function getAllPopulationServed(): Promise<CodeList[]> {
	try {
		const populationServed = await fetchAndParseCodeList(
			"https://codelist.commonapproach.org/PopulationServed/PopulationServed.owl"
		);

		return populationServed;
	} catch (error) {
		console.error("Error fetching PopulationServed code list:", error);
		return [];
	}
}

export async function getAllProvinceTerritory(): Promise<CodeList[]> {
	try {
		const provinceTerritory = await fetchAndParseCodeList(
			"https://codelist.commonapproach.org/ProvinceTerritory/ProvinceTerritory.owl"
		);

		return provinceTerritory;
	} catch (error) {
		console.error("Error fetching ProvinceTerritory code list:", error);
		return [];
	}
}

export async function getAllOrganizationType(): Promise<CodeList[]> {
	try {
		const organizationType = await fetchAndParseCodeList(
			"https://codelist.commonapproach.org/OrgTypeGOC/OrgTypeGOC.owl"
		);

		return organizationType;
	} catch (error) {
		console.error("Error fetching OrganizationType code list:", error);
		return [];
	}
}

export async function getAllLocalities(): Promise<CodeList[]> {
	try {
		const localities = await fetchAndParseCodeList(
			"https://codelist.commonapproach.org/Locality/LocalityStatsCan.owl"
		);

		return localities;
	} catch (error) {
		console.error("Error fetching Locality code list:", error);
		return [];
	}
}

export async function getAllCorporateRegistrars(): Promise<CodeList[]> {
	try {
		const corporateRegistrars = await fetchAndParseCodeList(
			"https://codelist.commonapproach.org/CanadianCorporateRegistries/CanadianCorporateRegistries.ttl"
		);

		return corporateRegistrars;
	} catch (error) {
		console.error("Error fetching CorporateRegistrar code list:", error);
		return [];
	}
}

export async function getAllIRISImpactCategories(): Promise<CodeList[]> {
	try {
		const irisCategories = await fetchAndParseCodeList(
			"https://codelist.commonapproach.org/IRISImpactThemes/IRISImpactCategories.ttl"
		);

		return irisCategories;
	} catch (error) {
		console.error("Error fetching IRIS+ Impact Categories code list:", error);
		return [];
	}
}

export async function getAllEquityDeservingGroups(): Promise<CodeList[]> {
	try {
		const edgGroups = await fetchAndParseCodeList(
			"https://codelist.commonapproach.org/EquityDeservingGroupsESDC/EquityDeservingGroupsESDC.ttl"
		);

		return edgGroups;
	} catch (error) {
		console.error("Error fetching Equity Deserving Groups code list:", error);
		return [];
	}
}