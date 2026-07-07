export function getSpreadsheetVersion(sheetNames: string[]): string {
  const ss = SpreadsheetApp.getActive();
  const updatedAt = (() => {
    try {
      return ss.getLastUpdated()?.getTime() || 0;
    } catch (_) {
      return 0;
    }
  })();
  const shape = sheetNames
    .map((name) => {
      try {
        const sh = ss.getSheetByName(name);
        if (!sh) return `${name}:missing`;
        return `${name}:${sh.getLastRow()}x${sh.getLastColumn()}`;
      } catch (_) {
        return `${name}:error`;
      }
    })
    .join('|');
  return `${updatedAt}:${shape}`;
}

export function readDocumentCachedJson<T>(options: {
  key: string;
  ttlSeconds: number;
  version: string;
  loader: () => T;
}): T {
  const { key, ttlSeconds, version, loader } = options;
  try {
    const cache = CacheService.getDocumentCache();
    const cached = cache.get(key);
    if (cached) {
      const parsed = JSON.parse(cached);
      if (parsed && parsed.version === version) {
        return parsed.value as T;
      }
    }
    const value = loader();
    try {
      cache.put(key, JSON.stringify({ version, value }), ttlSeconds);
    } catch (_) {
      // Ignore cache write failures and return the fresh value.
    }
    return value;
  } catch (_) {
    return loader();
  }
}

export function removeDocumentCacheKeys(keys: string[]) {
  if (!Array.isArray(keys) || !keys.length) return;
  try {
    const cache = CacheService.getDocumentCache();
    keys.filter(Boolean).forEach((key) => {
      try {
        cache.remove(key);
      } catch (_) {
        // Ignore per-key removal failures.
      }
    });
  } catch (_) {
    // Ignore cache removal failures.
  }
}
