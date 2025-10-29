// src/rpc.ts
import { getFilesForFolderUrl } from './util/drive';
import { addService, getServicePeople, esvPassage } from './features/services';

export function rpc(input: { method: string; payload: unknown }) {
  const { method, payload } = input || ({} as any);
  switch (method) {
    case 'getFilesForFolderUrl':
      return getFilesForFolderUrl(String(payload), 200);
    case 'addService':
      return addService(payload as any);
    case 'getServicePeople':
      return getServicePeople();
    case 'esvPassage':
      return esvPassage(payload as any);
    default:
      throw new Error(`Unknown RPC method: ${method}`);
  }
}
