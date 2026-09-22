// src/rpc.ts
import { getFilesForFolderUrl } from './util/drive';
import { getSongFolderUrl } from './features/song-media';
import { addService, createServicesBatch, getScriptureVersions, getServicePeople, esvPassage, getService, listServices, saveService, deleteService, getSongPerformances, suggestYouTubeStream, saveSongPerformanceTimestamp } from './features/services';
import { getServiceViewerStartup } from './features/service-viewer';
import { getOrder, getOrderRecordingLinks, saveOrder } from './features/order';
import { exportOrderOfWorshipDoc } from './features/order-of-worship-doc';
import { suggestSongs, getSongsForServiceView, getSongsWithLinksForView, rebuildSongUsageFromPlanner, getSongFields, updateSongRecency, saveSongEntry, suggestSongMetadata } from './features/songs';
import { aiScripturesForLyrics } from './util/ai';
import { listRoles, updateRoleEntry, addRoleEntry, memberExistsInRoles, getViewerProfile, getViewerAuthDebug } from './features/roles';
import { listWeeklyTeams, createWeeklyTeam, saveWeeklyTeam, saveWeeklyTeamDefaults } from './features/weekly-teams';
import { getTeamScheduleSnapshot, getServiceTeamAssignments, getUnavailableByServices, resetServiceTeamAssignments, saveServiceTeamAssignments } from './features/service-team-assignments';
import { getMemberAvailability, saveMemberAvailability } from './features/member-availability';
import { sendAvailabilityEmail, sendServiceTeamEmail } from './features/messaging';
import { summarizePassageWithSongs } from './features/scripture';
import { getVolunteerRequestsSnapshot, setViewerVolunteerRequest } from './features/volunteer-requests';

export function rpc(input: { method: string; payload: unknown }) {
  const { method, payload } = input || ({} as any);
  const startedAt = Date.now();
  try {
    switch (method) {
      case 'getFilesForFolderUrl':
        // The viewer needs a practical list of charts/tracks, not an
        // unbounded archive dump. Keeping this bounded avoids long-running
        // Drive metadata reads that can exceed the proxy's upstream window.
        return getFilesForFolderUrl(String(payload), 60);
      case 'getSongFolderUrl':
        return getSongFolderUrl(payload as { songName?: string });
      case 'addService':
        return addService(payload as any);
      case 'saveService':
        return saveService(payload as any);
      case 'deleteService':
        return deleteService(payload as any);
      case 'listServices':
        return listServices(payload as any);
      case 'getService':
        return getService(String(payload || ''));
      case 'getServiceViewerStartup':
        return getServiceViewerStartup();
      case 'createServicesBatch':
        return createServicesBatch(payload as any);
      case 'getOrder':
        return getOrder(String(payload || ''));
      case 'getOrderRecordingLinks':
        return getOrderRecordingLinks(payload as { songNames?: string[] });
      case 'saveOrder':
        return saveOrder(payload as any);
      case 'exportOrderOfWorshipDoc':
        return exportOrderOfWorshipDoc(payload as any);
      case 'suggestSongs':
        return suggestSongs(payload as any);
      case 'getSongsForView': return getSongsWithLinksForView();
      case 'getSongsForServiceView': return getSongsForServiceView(payload as { names?: string[] });
      case 'getSongPerformances': return getSongPerformances(payload as any);
      case 'saveSongPerformanceTimestamp': return saveSongPerformanceTimestamp(payload as any);
      case 'suggestYouTubeStream': return suggestYouTubeStream(payload as any);
      case 'getSongFields': return getSongFields(payload as any);
      case 'updateSongUsage': return updateSongRecency(payload as any);
      case 'saveSongEntry':
        return saveSongEntry(payload as any);
      case 'suggestSongMetadata':
        return suggestSongMetadata(payload as any);
      case 'aiScripturesForLyrics':
        return aiScripturesForLyrics(payload as any);
      case 'summarizeScriptureThemes':
        return summarizePassageWithSongs(payload as any);
      case 'rebuildSongUsage': return rebuildSongUsageFromPlanner();
      case 'getServicePeople':
        return getServicePeople();
      case 'esvPassage':
        return esvPassage(payload as any);
      case 'getScriptureVersions':
        return getScriptureVersions(payload as any);
      case 'listRoles':
        return listRoles();
      case 'getTeamStartup':
        return {
          roles: listRoles(),
          weeklyTeams: listWeeklyTeams()
        };
      case 'updateRoleEntry':
        return updateRoleEntry(payload as any);
      case 'addRoleEntry':
        return addRoleEntry(payload as any);
      case 'memberExistsInRoles':
        return memberExistsInRoles(payload as any);
      case 'getViewerProfile':
        return getViewerProfile();
      case 'getViewerAuthDebug':
        return getViewerAuthDebug();
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
      case 'getUnavailableByServices':
        return getUnavailableByServices(payload as any);
      case 'getServiceTeamAssignments':
        return getServiceTeamAssignments(payload as any);
      case 'resetServiceTeamAssignments':
        return resetServiceTeamAssignments(payload as any);
      case 'saveServiceTeamAssignments':
        return saveServiceTeamAssignments(payload as any);
      case 'getMemberAvailability':
        return getMemberAvailability(payload as any);
      case 'saveMemberAvailability':
        return saveMemberAvailability(payload as any);
      case 'getVolunteerRequestsSnapshot':
        return getVolunteerRequestsSnapshot(payload as any);
      case 'setViewerVolunteerRequest':
        return setViewerVolunteerRequest(payload as any);
      case 'sendAvailabilityEmail':
        return sendAvailabilityEmail(payload as any);
      case 'sendServiceTeamEmail':
        return sendServiceTeamEmail(payload as any);
      default:
        throw new Error(`Unknown RPC method: ${method}`);
    }
  } catch (err) {
    try { Logger.log(`RPC error (${method}): ${err && (err as any).stack || err}`); } catch(_) {}
    // Rethrow a clean error message so client failure handler triggers
    const msg = (err && (err as any).message) ? (err as any).message : String(err);
    throw new Error(msg);
  } finally {
    const elapsedMs = Date.now() - startedAt;
    // Keep routine execution logs clean while making slow sheet/proxy calls
    // visible in Apps Script Executions during performance investigations.
    if (elapsedMs >= 250) {
      try { Logger.log(`RPC timing method=${method} elapsedMs=${elapsedMs}`); } catch (_) { /* ignore */ }
    }
  }
}
