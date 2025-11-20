// src/rpc.ts
import { getFilesForFolderUrl } from './util/drive';
import { addService, createServicesBatch, getServicePeople, esvPassage, listServices, saveService, deleteService } from './features/services';
import { getOrder, saveOrder } from './features/order';
import { suggestSongs, getSongsWithLinksForView, rebuildSongUsageFromPlanner, getSongFields, updateSongRecency, saveSongEntry } from './features/songs';
import { aiScripturesForLyrics } from './util/ai';
import { listRoles, updateRoleEntry, addRoleEntry, memberExistsInRoles, getViewerProfile } from './features/roles';
import { listWeeklyTeams, createWeeklyTeam, saveWeeklyTeam, saveWeeklyTeamDefaults } from './features/weekly-teams';
import { getTeamScheduleSnapshot, getServiceTeamAssignments, saveServiceTeamAssignments } from './features/service-team-assignments';
import { getMemberAvailability, saveMemberAvailability } from './features/member-availability';
import { sendAvailabilityEmail } from './features/messaging';
import { summarizePassageWithSongs } from './features/scripture';

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
        return listServices(payload as any);
      case 'createServicesBatch':
        return createServicesBatch(payload as any);
      case 'getOrder':
        return getOrder(String(payload || ''));
      case 'saveOrder':
        return saveOrder(payload as any);
      case 'suggestSongs':
        return suggestSongs(payload as any);
      case 'getSongsForView': return getSongsWithLinksForView();
      case 'getSongFields': return getSongFields(payload as any);
      case 'updateSongUsage': return updateSongRecency(payload as any);
      case 'saveSongEntry':
        return saveSongEntry(payload as any);
      case 'aiScripturesForLyrics':
        return aiScripturesForLyrics(payload as any);
      case 'summarizeScriptureThemes':
        return summarizePassageWithSongs(payload as any);
      case 'rebuildSongUsage': return rebuildSongUsageFromPlanner();
      case 'getServicePeople':
        return getServicePeople();
      case 'esvPassage':
        return esvPassage(payload as any);
      case 'listRoles':
        return listRoles();
      case 'updateRoleEntry':
        return updateRoleEntry(payload as any);
      case 'addRoleEntry':
        return addRoleEntry(payload as any);
      case 'memberExistsInRoles':
        return memberExistsInRoles(payload as any);
      case 'getViewerProfile':
        return getViewerProfile();
      case 'listWeeklyTeams':
        return listWeeklyTeams();
      case 'createWeeklyTeam':
        return createWeeklyTeam(payload as any);
      case 'saveWeeklyTeam':
        return saveWeeklyTeam(payload as any);
      case 'saveWeeklyTeamDefaults':
        return saveWeeklyTeamDefaults(payload as any);
      case 'getTeamScheduleSnapshot':
        return getTeamScheduleSnapshot(payload as any);
      case 'getServiceTeamAssignments':
        return getServiceTeamAssignments(payload as any);
      case 'saveServiceTeamAssignments':
        return saveServiceTeamAssignments(payload as any);
      case 'getMemberAvailability':
        return getMemberAvailability(payload as any);
      case 'saveMemberAvailability':
        return saveMemberAvailability(payload as any);
      case 'sendAvailabilityEmail':
        return sendAvailabilityEmail(payload as any);
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
