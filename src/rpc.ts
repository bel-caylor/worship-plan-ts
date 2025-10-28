// src/rpc.ts
import { getFilesForFolderUrl } from './util/drive';

export function rpc(input: { method: string; payload: unknown }) {
  const { method, payload } = input || ({} as any);
  switch (method) {
    case 'getFilesForFolderUrl':
      return getFilesForFolderUrl(String(payload), 200);
    default:
      throw new Error(`Unknown RPC method: ${method}`);
  }
}
