// Utilities for working with JSON-LD @type values which may be a string or an array
// and may include multiple namespace-prefixed entries. We prioritize cids: and sff: types.

export function getPrimaryStandardType(typeVal: any): string | null {
  if (!typeVal) return null;
  const isTarget = (t: string) => 
    typeof t === "string" && 
    (t.startsWith("cids:") || t.startsWith("sff:") || t.startsWith("org:"));
  if (typeof typeVal === "string") return isTarget(typeVal) ? typeVal : null;
  if (Array.isArray(typeVal)) {
    const found = typeVal.find(isTarget);
    return found || null;
  }
  return null;
}

export function getCidsTableSuffix(typeVal: any): string | null {
  const main = getPrimaryStandardType(typeVal);
  if (main) return main.split(":")[1];
  if (typeof typeVal === "string") {
    return typeVal.includes(":") ? typeVal.split(":")[1] : typeVal;
  }
  if (Array.isArray(typeVal)) {
    const first = typeVal.find((t) => typeof t === "string");
    if (!first) return null;
    return first.includes(":") ? first.split(":")[1] : first;
  }
  return null;
}
