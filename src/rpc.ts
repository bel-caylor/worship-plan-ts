// src/rpc.ts
import { getFilesForFolderUrl } from './util/drive';
import { addService, getServicePeople, esvPassage, listServices, saveService } from './features/services';
import { getOrder, saveOrder } from './features/order';

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
      case 'listServices':
        return listServices();
      case 'getOrder':
        return getOrder(String(payload || ''));
      case 'saveOrder':
        return saveOrder(payload as any);
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
