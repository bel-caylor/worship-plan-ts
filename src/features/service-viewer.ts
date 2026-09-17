import { getOrder } from './order';
import { getNextUpcomingService } from './services';

/**
 * The public viewer's critical-path payload. It intentionally contains only
 * the next service and its saved order; service history and song details load
 * afterward without delaying the first useful screen.
 */
export function getServiceViewerStartup() {
  const service = getNextUpcomingService();
  if (!service?.id) return { service: null, items: [] };
  return {
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
}
