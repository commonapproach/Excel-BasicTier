/* global fetch, console, localStorage */
// Fetch, cache, and parse SELI-GLI Turtle data for Theme, Outcome, Indicator
export interface SeliTheme {
  "@id": string;
  hasName: string;
}
export interface SeliOutcome {
  "@id": string;
  hasName: string;
  forTheme: string; // @id of Theme
  hasIndicator: string[]; // @ids of Indicators
}
export interface SeliIndicator {
  "@id": string;
  hasName: string;
  hasDescription?: string;
  forOutcome: string; // @id of Outcome
}

export interface SeliGLIData {
  themes: SeliTheme[];
  outcomes: SeliOutcome[];
  indicators: SeliIndicator[];
}

const SELI_GLI_URL = "https://codelist.commonapproach.org/SELI-GLI.ttl";
const SELI_GLI_GITHUB_FALLBACK_URL = "https://raw.githubusercontent.com/commonapproach/CodeLists/main/SELI-GLI.ttl";
const CACHE_KEY = "seli_gli_cache";
const CACHE_EXPIRATION = 24 * 60 * 60 * 1000;

// Internal debug logger (disabled by default to satisfy no-console lint rule)
const __DEBUG_SELI = false;
function dLog(...args: any[]) {
  if (__DEBUG_SELI) {
    // eslint-disable-next-line no-console
    console.log(...args);
  }
}
function dWarn(...args: any[]) {
  if (__DEBUG_SELI) {
    // eslint-disable-next-line no-console
    console.warn(...args);
  }
}
function dError(...args: any[]) {
  if (__DEBUG_SELI) {
    // eslint-disable-next-line no-console
    console.error(...args);
  }
}

function parseTurtleToSeliGLI(ttl: string): SeliGLIData {
    // New format uses @base and @prefix : <#>
    // @base <https://codelist.commonapproach.org/SELI-GLI> .
    // :Theme1 resolves to https://codelist.commonapproach.org/SELI-GLI#Theme1
    const baseUriMatch = ttl.match(/@base\s*<([^>]+)>/);
    const baseUri = baseUriMatch ? baseUriMatch[1] : "https://codelist.commonapproach.org/SELI-GLI";
    const fragmentBase = baseUri + "#";

    const themes: SeliTheme[] = [];
    const outcomes: SeliOutcome[] = [];
    const indicators: SeliIndicator[] = [];

    // Split into subject blocks — each starts with :IdentifierN
    // Blocks are separated by blank lines or next :Identifier
    const blockRegex = /^(:[\w]+)\s*\n([\s\S]*?)(?=\n^:[\w]|\Z)/gm;
    // Simpler approach: split on lines that start with ":"
    const blocks = ttl.split(/\n(?=:[\w]+\s*\n\s+a\s+cids:)/);

    for (const block of blocks) {
        // Get subject id
        const subjectMatch = block.match(/^:([\w]+)/);
        if (!subjectMatch) continue;
        const localId = subjectMatch[1];
        const fullId = fragmentBase + localId;

        // Determine type
        const typeMatch = block.match(/a\s+cids:(Theme|Outcome|Indicator)[,\s;]/);
        if (!typeMatch) continue;
        const type = typeMatch[1];

        // Extract org:hasName (new format)
        const nameMatch = block.match(/org:hasName\s+"([^"]+)"/);
        if (!nameMatch) continue;
        const hasName = nameMatch[1];

        // Extract cids:hasDescription
        const descMatch = block.match(/cids:hasDescription\s+"([^"]+)"/);
        const hasDescription = descMatch ? descMatch[1] : undefined;

        if (type === "Theme") {
            themes.push({ "@id": fullId, hasName });
        } else if (type === "Outcome") {
            // Extract forTheme
            const forThemeMatch = block.match(/cids:forTheme\s+:([\w]+)/);
            const forTheme = forThemeMatch ? fragmentBase + forThemeMatch[1] : "";

            // Extract hasIndicator (comma-separated list)
            const hasIndicatorMatch = block.match(/cids:hasIndicator\s+((?::([\w]+)(?:\s*,\s*)?)+)/);
            let hasIndicator: string[] = [];
            if (hasIndicatorMatch) {
                hasIndicator = [...hasIndicatorMatch[1].matchAll(/:([\w]+)/g)]
                    .map((m) => fragmentBase + m[1]);
            }

            outcomes.push({ "@id": fullId, hasName, forTheme, hasIndicator });
        } else if (type === "Indicator") {
            // Extract forOutcome
            const forOutcomeMatch = block.match(/cids:forOutcome\s+:([\w]+)/);
            const forOutcome = forOutcomeMatch ? fragmentBase + forOutcomeMatch[1] : "";

            indicators.push({ "@id": fullId, hasName, hasDescription, forOutcome });
        }
    }

    return { themes, outcomes, indicators };
}

export async function fetchAndParseSeliGLI(): Promise<SeliGLIData> {
  // Check cache
  const cached = localStorage.getItem(CACHE_KEY);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (parsed.timestamp && Date.now() - parsed.timestamp < CACHE_EXPIRATION) {
        return parsed.data;
      }
    } catch {
      // Invalid cache, ignore
    }
    localStorage.removeItem(CACHE_KEY);
  }

  let ttl: string;
  let data: SeliGLIData;

  // Try to fetch from primary URL first
  try {
    dLog(`Attempting to fetch SELI-GLI from primary URL: ${SELI_GLI_URL}`);
    const resp = await fetch(SELI_GLI_URL);

    if (!resp.ok) {
      throw new Error(`Primary fetch failed with status: ${resp.status}`);
    }

    ttl = await resp.text();
    data = parseTurtleToSeliGLI(ttl);

    dLog(
      `Successfully fetched SELI-GLI data from primary URL - Themes: ${data.themes.length}, Outcomes: ${data.outcomes.length}, Indicators: ${data.indicators.length}`
    );
  } catch (primaryError) {
    dWarn("Primary SELI-GLI fetch failed:", primaryError);

    // Try GitHub fallback
    try {
      dLog(`Attempting SELI-GLI fallback from GitHub: ${SELI_GLI_GITHUB_FALLBACK_URL}`);
      const fallbackResp = await fetch(SELI_GLI_GITHUB_FALLBACK_URL);

      if (!fallbackResp.ok) {
        throw new Error(`Fallback fetch failed with status: ${fallbackResp.status}`);
      }

      ttl = await fallbackResp.text();
      data = parseTurtleToSeliGLI(ttl);

      dLog(
        `Successfully fetched SELI-GLI data from GitHub fallback - Themes: ${data.themes.length}, Outcomes: ${data.outcomes.length}, Indicators: ${data.indicators.length}`
      );
    } catch (fallbackError) {
      dError("Both primary and fallback SELI-GLI fetch failed:", fallbackError);

      // As a last resort, check if we have any cached data (even if expired)
      const lastResortCache = localStorage.getItem(CACHE_KEY);
      if (lastResortCache) {
        try {
          const parsedData = JSON.parse(lastResortCache);
          if (parsedData.data) {
            dWarn("Using expired cached SELI-GLI data as fallback");
            return parsedData.data;
          }
        } catch (cacheError) {
          dError("Failed to parse last resort SELI-GLI cache:", cacheError);
        }
      }

      throw new Error(
        `All SELI-GLI fetch attempts failed. Primary: ${
          primaryError instanceof Error ? primaryError.message : String(primaryError)
        }, Fallback: ${
          (fallbackError as any) instanceof Error
            ? (fallbackError as Error).message
            : String(fallbackError)
        }`
      );
    }
  }

  // Cache the successful data
  localStorage.setItem(CACHE_KEY, JSON.stringify({ timestamp: Date.now(), data }));
  return data;
}
