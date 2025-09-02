/* global fetch, console */
// Lightweight optional debug logger (disabled by default to satisfy no-console rule)
const __DEBUG_CONTEXT = false;
function ctxDebugError(...args: any[]) {
  if (__DEBUG_CONTEXT) {
    // eslint-disable-next-line no-console
    console.error(...args);
  }
}

interface CacheEntry {
  data: unknown;
  timestamp: number;
}

const cache = new Map<string, CacheEntry>();
const CACHE_TTL = 24 * 60 * 60 * 1000; // 24 hours in milliseconds

export async function getContext(url: string, ttl: number = CACHE_TTL): Promise<unknown> {
  const now = Date.now();
  const cached = cache.get(url);

  // Return cached data if it's still valid
  if (cached && now - cached.timestamp < ttl) {
    return cached.data;
  }

  try {
    const response = await fetch(url);
    const data = await response.json();

    // Store in cache
    cache.set(url, { data, timestamp: now });

    return data;
  } catch (error) {
    ctxDebugError(`Error fetching or parsing ${url}:`, error);
    throw error;
  }
}
