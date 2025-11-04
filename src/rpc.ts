// src/rpc.ts
import { getFilesForFolderUrl } from './util/drive';
import { addService, getServicePeople, esvPassage, listServices, saveService, deleteService } from './features/services';
import { getOrder, saveOrder } from './features/order';
import { suggestSongs, getSongsWithLinksForView, rebuildSongUsageFromPlanner, getSongFields } from './features/songs';
import { aiScripturesForLyrics } from './util/ai';

export function rpc(input: { method: string; payload: unknown }) {
  const { method, payload } = input || ({} as any);
  try {
    switch (method) {
      case 'getFilesForFolderUrl':
        return getFilesForFolderUrl(String(payload), 200);
      case 'addService':
        return addService(payload as any);
      case 'saveService':
        return saveService(payload as any);
      case 'deleteService':
        return deleteService(payload as any);
      case 'listServices':
        return listServices();
      case 'getOrder':
        return getOrder(String(payload || ''));
      case 'saveOrder':
        return saveOrder(payload as any);
      case 'suggestSongs':
        return suggestSongs(payload as any);
      case 'getSongsForView': return getSongsWithLinksForView();
      case 'getSongFields': return getSongFields(payload as any);
      case 'aiScripturesForLyrics':
        return aiScripturesForLyrics(payload as any);
      case 'rebuildSongUsage': return rebuildSongUsageFromPlanner();
      case 'getServicePeople':
        return getServicePeople();
      case 'esvPassage':
        return esvPassage(payload as any);
      default:
        throw new Error(`Unknown RPC method: ${method}`);
    }
  } catch (err) {
    try { Logger.log(`RPC error (${method}): ${err && (err as any).stack || err}`); } catch(_) {}
    // Rethrow a clean error message so client failure handler triggers
    const msg = (err && (err as any).message) ? (err as any).message : String(err);
    throw new Error(msg);
  }
}

