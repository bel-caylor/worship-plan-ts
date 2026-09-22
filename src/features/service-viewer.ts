import { getOrder } from './order';
import { getNextUpcomingServiceSummary } from './services';

const SERVICE_VIEWER_STARTUP_CACHE_KEY = 'serviceViewer:startup:v1';
const SERVICE_VIEWER_STARTUP_CACHE_TTL_SECONDS = 120;

function readStartupCache() {
  try {
    const raw = CacheService.getDocumentCache().get(SERVICE_VIEWER_STARTUP_CACHE_KEY);
    const parsed = raw ? JSON.parse(raw) : null;
    if (parsed && typeof parsed === 'object' && 'service' in parsed && Array.isArray(parsed.items)) {
      return parsed;
    }
  } catch (_) {
    // Cache misses should never affect the public viewer.
  }
  return null;
}

function writeStartupCache(value: unknown) {
  try {
    CacheService.getDocumentCache().put(
      SERVICE_VIEWER_STARTUP_CACHE_KEY,
      JSON.stringify(value),
      SERVICE_VIEWER_STARTUP_CACHE_TTL_SECONDS
    );
  } catch (_) {
    // Ignore cache capacity/serialization failures.
  }
}

/**
 * The public viewer's critical-path payload. It intentionally contains only
 * the next service and its saved order; service history and song details load
 * afterward without delaying the first useful screen.
 */
export function getServiceViewerStartup() {
  const cached = readStartupCache();
  if (cached) return cached;
  const service = getNextUpcomingServiceSummary();
  if (!service?.id) return { service: null, items: [] };
  const payload = {
    // The viewer heading needs only these fields. In particular, do not send
    // saved scripture text or suggested-song JSON on its critical path.
    service: {
      id: service.id,
      date: service.date,
      time: service.time,
      type: service.type,
      leader: service.leader
    },
    items: getOrder(service.id).items
  };
  writeStartupCache(payload);
  return payload;
}
