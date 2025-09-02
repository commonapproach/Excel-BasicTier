/* global fetch, localStorage */
// Fetch and parse Units of Measure TTL (or use a safe fallback)
// Provides both UI options (label + IRI) and a type/category for export logic.

export type UnitOption = { id: string; name: string };

export type UnitCategory =
  | "count"
  | "percent"
  | "ratio"
  | "score"
  | "ambiguous" // SI and monetary where Sum/Mean cannot be inferred
  | "unspecified";

// Canonical IRIs used by our logic (stable regardless of label localization)
export const UNIT_IRI = {
  // Use cids:countUnit (owl:sameAs i72:population_cardinality_unit)
  COUNT: "https://ontology.commonapproach.org/cids#countUnit",
  PERCENT: "https://ontology.commonapproach.org/cids#percentUnit",
  RATIO: "https://ontology.commonapproach.org/cids#ratioUnit",
  SCORE: "https://ontology.commonapproach.org/cids#scoreUnit",
  CAD: "https://ontology.commonapproach.org/cids#cad",
  CAD_MILLIONS: "https://ontology.commonapproach.org/cids#cadMillions",
  DAY: "https://ontology.commonapproach.org/cids#dayUnit",
  KILOGRAM: "https://ontology.commonapproach.org/cids#kilogram",
  KWH: "https://ontology.commonapproach.org/cids#kilowattHour",
  TONNE: "https://ontology.commonapproach.org/cids#tonne",
  MEGATONNE: "https://ontology.commonapproach.org/cids#megatonne",
  MEGAWATT: "https://ontology.commonapproach.org/cids#megawatt",
  MEGAWATT_HOUR: "https://ontology.commonapproach.org/cids#megawattHour",
  SQUARE_METER: "https://ontology.commonapproach.org/cids#squareMeter",
  CUBIC_METER: "https://ontology.commonapproach.org/cids#cubicMeter",
  SQUARE_FOOT: "https://ontology.commonapproach.org/cids#squareFoot",
  FOOT: "https://ontology.commonapproach.org/cids#foot",
  UNSPECIFIED: "https://ontology.commonapproach.org/cids#unspecifiedUnit",
} as const;

// Minimal fallback set aligned with customer table and the draft TTL.
const FALLBACK_UNITS: { id: string; name: string; category: UnitCategory }[] = [
  { id: UNIT_IRI.COUNT, name: "Count", category: "count" },
  { id: UNIT_IRI.PERCENT, name: "Percent", category: "percent" },
  { id: UNIT_IRI.RATIO, name: "Ratio", category: "ratio" },
  { id: UNIT_IRI.SCORE, name: "Score", category: "score" },
  { id: UNIT_IRI.CAD, name: "Canadian Dollar", category: "ambiguous" },
  { id: UNIT_IRI.CAD_MILLIONS, name: "Millions of Canadian Dollars", category: "ambiguous" },
  { id: UNIT_IRI.DAY, name: "Days", category: "ambiguous" },
  { id: UNIT_IRI.KILOGRAM, name: "Kilogram", category: "ambiguous" },
  { id: UNIT_IRI.KWH, name: "Kilowatt-hour", category: "ambiguous" },
  { id: UNIT_IRI.TONNE, name: "Tonne", category: "ambiguous" },
  { id: UNIT_IRI.MEGATONNE, name: "Megatonne", category: "ambiguous" },
  { id: UNIT_IRI.MEGAWATT, name: "Megawatt", category: "ambiguous" },
  { id: UNIT_IRI.MEGAWATT_HOUR, name: "Megawatt-hour", category: "ambiguous" },
  { id: UNIT_IRI.SQUARE_METER, name: "Square Meter", category: "ambiguous" },
  { id: UNIT_IRI.CUBIC_METER, name: "Cubic Meter", category: "ambiguous" },
  { id: UNIT_IRI.SQUARE_FOOT, name: "Square Foot", category: "ambiguous" },
  { id: UNIT_IRI.FOOT, name: "Foot", category: "ambiguous" },
  { id: UNIT_IRI.UNSPECIFIED, name: "Unspecified Unit", category: "unspecified" },
];

// Category lookup by IRI for quick business-logic decisions
export const UNIT_CATEGORY_BY_IRI: Record<string, UnitCategory> = Object.fromEntries(
  FALLBACK_UNITS.map((u) => [u.id, u.category])
);

// Internal cache for fetched catalogs
let CACHED_OPTIONS: UnitOption[] | null = null;
let CACHED_DEFINITIONS: Record<string, any> | null = null;

// Persisted cache (24h) to avoid re-fetching on every session
const CACHE_KEY = "units_of_measure_cache";
const CACHE_EXPIRATION = 24 * 60 * 60 * 1000; // 24 hours

type UnitCatalogCache = {
  timestamp: number;
  options: UnitOption[];
  definitions: Record<string, any>;
};

function loadFromLocalStorageCache(): boolean {
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (!raw) return false;
    const parsed = JSON.parse(raw) as UnitCatalogCache;
    if (!parsed?.timestamp || Date.now() - parsed.timestamp >= CACHE_EXPIRATION) {
      localStorage.removeItem(CACHE_KEY);
      return false;
    }
    CACHED_OPTIONS = parsed.options;
    CACHED_DEFINITIONS = parsed.definitions;
    return Array.isArray(CACHED_OPTIONS) && CACHED_OPTIONS.length > 0;
  } catch {
    // ignore invalid cache
    try {
      localStorage.removeItem(CACHE_KEY);
    } catch {
      // no-op
    }
    return false;
  }
}

function saveToLocalStorageCache(options: UnitOption[], definitions: Record<string, any>): void {
  try {
    const payload: UnitCatalogCache = {
      timestamp: Date.now(),
      options,
      definitions,
    };
    localStorage.setItem(CACHE_KEY, JSON.stringify(payload));
  } catch {
    // storage might be unavailable or quota-exceeded; ignore
  }
}

// Public API: options for UI select (id = IRI, name = label), prioritized by server -> GitHub -> fallback
export async function getUnitOptions(): Promise<UnitOption[]> {
  await ensureCatalogLoaded();
  return CACHED_OPTIONS && CACHED_OPTIONS.length > 0
    ? CACHED_OPTIONS
    : FALLBACK_UNITS.map(({ id, name }) => ({ id, name }));
}

export async function getUnitDefinition(iri: string): Promise<any | null> {
  await ensureCatalogLoaded();
  return (CACHED_DEFINITIONS && CACHED_DEFINITIONS[iri]) || UNIT_DEFINITIONS[iri] || null;
}

async function ensureCatalogLoaded(): Promise<void> {
  if (CACHED_OPTIONS && CACHED_DEFINITIONS) return;
  // Try persisted cache first
  if (loadFromLocalStorageCache()) return;
  const ttlUrlCandidates = [
    "https://codelist.commonapproach.org/codeLists/UnitsOfMeasureList.ttl",
    "https://raw.githubusercontent.com/commonapproach/CodeLists/main/UnitOfMeasure/UnitsOfMeasureList.ttl",
  ];
  for (const url of ttlUrlCandidates) {
    try {
      const resp = await fetch(url, { cache: "no-cache" });
      if (!resp.ok) continue;
      const ttl = await resp.text();
      const { options, definitions } = parseTtlToCatalog(ttl);
      // Only keep cids namespace
      const filteredOptions = options.filter((o) =>
        o.id.startsWith("https://ontology.commonapproach.org/cids#")
      );
      const filteredDefinitions = Object.fromEntries(
        Object.entries(definitions).filter(([iri]) =>
          iri.startsWith("https://ontology.commonapproach.org/cids#")
        )
      );
      if (filteredOptions.length > 0) {
        CACHED_OPTIONS = filteredOptions;
        CACHED_DEFINITIONS = filteredDefinitions;
        saveToLocalStorageCache(CACHED_OPTIONS, CACHED_DEFINITIONS);
        return;
      }
    } catch (_) {
      // try next source
    }
  }
  // As a last resort, populate cache from fallback so subsequent calls are consistent
  CACHED_OPTIONS = FALLBACK_UNITS.map(({ id, name }) => ({ id, name }));
  CACHED_DEFINITIONS = { ...UNIT_DEFINITIONS };
  saveToLocalStorageCache(CACHED_OPTIONS, CACHED_DEFINITIONS);
}

// Helper: parse a minimal subset of TTL for cids:* terms into options and definitions
function parseTtlToCatalog(ttl: string): {
  options: UnitOption[];
  definitions: Record<string, any>;
} {
  // Capture prefix mapping like: @prefix cids: <...> .
  const prefixMap: Record<string, string> = {};
  const prefixRegex = /@prefix\s+([a-zA-Z][\w-]*):\s*<([^>]+)>\s*\./g;
  let pm: RegExpExecArray | null;
  while ((pm = prefixRegex.exec(ttl))) {
    prefixMap[pm[1]] = pm[2];
  }

  const options: UnitOption[] = [];
  const definitions: Record<string, any> = {};
  // Match blocks that contain rdfs:label. We'll also search for rdf:type and i72:symbol within the same block.
  const blockRegex =
    /(\w+):(\w+)\s+[\s\S]*?rdfs:label\s+"([^"]+)"@en\s*;[\s\S]*?(?=\n\w+:\w+\s|$)/g;
  let m: RegExpExecArray | null;
  while ((m = blockRegex.exec(ttl))) {
    const pfx = m[1];
    const local = m[2];
    const label = m[3];
    const base = prefixMap[pfx];
    if (!base) continue;
    const subjectIri = base + local;
    options.push({ id: subjectIri, name: label });

    const block = m[0];
    const typeMatches = [...block.matchAll(/rdf:type\s+(\w+:\w+)/g)].map((tm) => tm[1]);
    const symbolMatch = block.match(/i72:symbol\s+"([^"]+)"@en\s*;/);

    const def: any = { "@id": subjectIri, "rdfs:label": label };
    if (typeMatches.length > 0) def["@type"] = typeMatches[0];
    if (symbolMatch) def["i72:symbol"] = symbolMatch[1];
    definitions[subjectIri] = def;
  }

  return { options, definitions };
}

// Export full JSON-LD definitions for known units so we can embed them on export
export const UNIT_DEFINITIONS: Record<string, any> = {
  [UNIT_IRI.COUNT]: {
    "@id": UNIT_IRI.COUNT,
    "@type": "i72:Cardinality_unit",
    "rdfs:label": "Count",
    "i72:symbol": "count",
    "owl:sameAs": "http://ontology.eil.utoronto.ca/ISO21972/iso21972#population_cardinality_unit",
  },
  [UNIT_IRI.PERCENT]: {
    "@id": UNIT_IRI.PERCENT,
    "@type": "i72:Singular_unit",
    "rdfs:label": "Percent",
    "i72:symbol": "%",
  },
  [UNIT_IRI.RATIO]: {
    "@id": UNIT_IRI.RATIO,
    "@type": "i72:Compound_unit",
    "rdfs:label": "Ratio",
    "i72:symbol": "ratio",
  },
  [UNIT_IRI.SCORE]: {
    "@id": UNIT_IRI.SCORE,
    "@type": "i72:Singular_unit",
    "rdfs:label": "Score",
    "i72:symbol": "score",
  },
  [UNIT_IRI.CAD]: {
    "@id": UNIT_IRI.CAD,
    "@type": "i72:Monetary_unit",
    "rdfs:label": "Canadian Dollar",
    "i72:symbol": "$CAD",
  },
  [UNIT_IRI.CAD_MILLIONS]: {
    "@id": UNIT_IRI.CAD_MILLIONS,
    "@type": "i72:Unit_multiple_or_submultiple",
    "rdfs:label": "Millions of Canadian Dollars",
    "i72:symbol": "$CAD(M)",
    "i72:prefix": "i72:mega",
    "i72:singular_unit": UNIT_IRI.CAD,
  },
  [UNIT_IRI.DAY]: {
    "@id": UNIT_IRI.DAY,
    "@type": "i72:Singular_unit",
    "rdfs:label": "Days",
    "i72:symbol": "days",
  },
  [UNIT_IRI.KILOGRAM]: {
    "@id": UNIT_IRI.KILOGRAM,
    "@type": "i72:Unit_multiple_or_submultiple",
    "rdfs:label": "Kilogram",
    "i72:symbol": "kg",
    "i72:prefix": "i72:kilo",
    "i72:singular_unit": "i72:gram",
  },
  [UNIT_IRI.KWH]: {
    "@id": UNIT_IRI.KWH,
    "@type": "i72:Unit_multiplication",
    "rdfs:label": "Kilowatt-hour",
    "i72:symbol": "kWh",
    "i72:term_1": "i72:kilowatt",
    "i72:term_2": "i72:hour",
  },
  [UNIT_IRI.TONNE]: {
    "@id": UNIT_IRI.TONNE,
    "@type": "i72:Singular_unit",
    "rdfs:label": "Tonne",
    "i72:symbol": "t",
  },
  [UNIT_IRI.MEGATONNE]: {
    "@id": UNIT_IRI.MEGATONNE,
    "@type": "i72:Unit_multiple_or_submultiple",
    "rdfs:label": "Megatonne",
    "i72:symbol": "Mt",
    "i72:prefix": "i72:mega",
    "i72:singular_unit": UNIT_IRI.TONNE,
  },
  [UNIT_IRI.MEGAWATT]: {
    "@id": UNIT_IRI.MEGAWATT,
    "@type": "i72:Unit_multiple_or_submultiple",
    "rdfs:label": "Megawatt",
    "i72:symbol": "MW",
    "i72:prefix": "i72:mega",
    "i72:singular_unit": "i72:watt",
  },
  [UNIT_IRI.MEGAWATT_HOUR]: {
    "@id": UNIT_IRI.MEGAWATT_HOUR,
    "@type": "i72:Unit_multiplication",
    "rdfs:label": "Megawatt-hour",
    "i72:symbol": "MWh",
    "i72:term_1": UNIT_IRI.MEGAWATT,
    "i72:term_2": "i72:hour",
  },
  [UNIT_IRI.SQUARE_METER]: {
    "@id": UNIT_IRI.SQUARE_METER,
    "@type": "i72:Unit_exponentiation",
    "rdfs:label": "Square Meter",
    "i72:symbol": "m^2",
    "i72:base": "i72:metre",
    "i72:exponent": 2,
  },
  [UNIT_IRI.CUBIC_METER]: {
    "@id": UNIT_IRI.CUBIC_METER,
    "@type": "i72:Unit_exponentiation",
    "rdfs:label": "Cubic Meter",
    "i72:symbol": "m^3",
    "i72:base": "i72:metre",
    "i72:exponent": 3,
  },
  [UNIT_IRI.SQUARE_FOOT]: {
    "@id": UNIT_IRI.SQUARE_FOOT,
    "@type": "i72:Unit_exponentiation",
    "rdfs:label": "Square Foot",
    "i72:symbol": "sq ft",
    "i72:base": UNIT_IRI.FOOT,
    "i72:exponent": 2,
  },
  [UNIT_IRI.FOOT]: {
    "@id": UNIT_IRI.FOOT,
    "@type": "i72:Singular_unit",
    "rdfs:label": "Foot",
    "i72:symbol": "ft",
  },
  [UNIT_IRI.UNSPECIFIED]: {
    "@id": UNIT_IRI.UNSPECIFIED,
    "@type": "i72:Singular_unit",
    "rdfs:label": "Unspecified Unit",
  },
};
