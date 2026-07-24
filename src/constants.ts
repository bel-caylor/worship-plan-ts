// src/constants.ts
export const SONG_SHEET = 'Songs';
export const SONG_COL_NAME = 'Song';
export const FOLDER_LINK_COL = 'Folder URL';
export const AUDIO_LINKS_COL = 'Audio Files';
export const MAX_AUDIO_LINKS = 5;
export const AUDIO_MIME_PREFIX = 'audio/';
export const AUDIO_EXT = new Set(['mp3','m4a','aac','wav','aiff','aif','flac','ogg','oga','opus','wma']);
export const ROOT_FOLDER_ID = '19buHshZq5phnvP8xvFHnwkTnwdvg0FV_';
export const SPANISH_ROOT_ID = '1bYk1utXCF0D1r5a_GHlAjqmO5gLBg8b5';
export const SP_COL_NAME = 'Sp';
export const TARGET_LEADER_COL = 'Leader';
export const PLANNER_SHEET = 'Weekly Planner';
export const PLANNER_LEADER_CANDIDATES = ['Leader'];
export const PLANNER_SONG_COLS = ['Opening Song','Song2','Song3','Song4/Communion','Offering/Communion Song','Closing Song'];
export const SERVICES_SHEET = 'Services';
export const SERVICES_COL = {
  id: 'ServiceID',
  date: 'Date',
  time: 'Time',
  type: 'ServiceType',
  youtubeUrl: 'YouTube URL',
  leader: 'Leader',
  preacher: 'Preacher',
  scripture: 'Scripture',
  scriptureText: 'Scripture Text',
  theme: 'Theme',
  keywords: 'Keywords',
  notes: 'Notes',
  suggestedSongs: 'Suggested Songs'
} as const;

export const YOUTUBE_STREAMS_SHEET = 'YouTube Streams';
export const YOUTUBE_STREAMS_COL = {
  videoId: 'VideoId',
  url: 'YouTube URL',
  title: 'Title',
  streamDate: 'Stream Date',
  publishedDate: 'Published Date',
  channelId: 'Channel Id',
  channelName: 'Channel Name',
  matchedServiceId: 'Matched Service ID',
  status: 'Status',
  notes: 'Notes',
  source: 'Source'
} as const;

export const SONG_PERFORMANCES_SHEET = 'SongPerformances';
export const SONG_PERFORMANCES_COL = {
  songId: 'SongId',
  songName: 'Song Name',
  serviceId: 'ServiceID',
  youtubeUrl: 'YouTube URL',
  videoId: 'VideoId',
  startSeconds: 'StartSeconds',
  startLabel: 'StartLabel',
  matchSource: 'MatchSource',
  matchConfidence: 'MatchConfidence',
  matchedLyric: 'MatchedLyric',
  notes: 'Notes',
  lastVerified: 'LastVerified'
} as const;

// Order of Worship sheet configuration
export const ORDER_SHEET = 'ServiceItems';
export const ORDER_COL = {
  serviceId: 'ServiceID',
  order: 'Order',
  itemType: 'ItemType',
  detail: 'Detail',
  scriptureText: 'ScriptureText',
  leader: 'Leader',
  notes: 'Notes'
} as const;

export const ORDER_OF_WORSHIP_EXPORT_FOLDER_URL = 'https://drive.google.com/drive/folders/13cYAQ22ntObng64BPcGJf7_R6uSvVZBn';

export type Row = Record<string, unknown>;

export const ROLES_SHEET = 'Roles';
export const ROLES_COL = {
  email: 'Email',
  permissions: 'Permissions',
  first: 'First',
  last: 'Last',
  team: 'Team',
  role: 'Role',
  spanish: 'Spanish'
} as const;

export const WEEKLY_TEAMS_SHEET = 'WeeklyTeams';
export const WEEKLY_TEAMS_COL = {
  team: 'Team',
  teamName: 'TeamName',
  description: 'Description'
} as const;

export const WEEKLY_TEAM_ROLES_SHEET = 'WeeklyTeamRoles';
export const WEEKLY_TEAM_ROLES_COL = {
  team: 'Team',
  teamName: 'TeamName',
  roleType: 'RoleType',
  roleName: 'Role',
  memberEmail: 'MemberEmail',
  memberName: 'MemberName'
} as const;

export const WEEKLY_TEAM_ROLE_DEFAULTS_SHEET = 'WeeklyTeamRoleDefaults';
export const WEEKLY_TEAM_ROLE_DEFAULTS_COL = {
  team: 'Team',
  roleName: 'RoleName',
  order: 'Order'
} as const;

export const MEMBER_AVAILABILITY_SHEET = 'MemberAvailability';
export const MEMBER_AVAILABILITY_COL = {
  email: 'Email',
  serviceId: 'ServiceID',
  availability: 'Availability'
} as const;

export const SERVICE_TEAM_ASSIGNMENTS_SHEET = 'ServiceTeamAssignments';
export const SERVICE_TEAM_ASSIGNMENTS_COL = {
  serviceId: 'ServiceID',
  serviceType: 'ServiceType',
  teamType: 'Team',
  weeklyTeamName: 'WeeklyTeam',
  roleName: 'Role',
  roleType: 'RoleType',
  memberEmail: 'MemberEmail',
  memberName: 'MemberName',
  status: 'Status',
  notes: 'Notes'
} as const;

export const VOLUNTEER_REQUESTS_SHEET = 'VolunteerRequests';
export const VOLUNTEER_REQUESTS_COL = {
  serviceId: 'ServiceID',
  teamType: 'Team',
  roleName: 'Role',
  memberEmail: 'MemberEmail',
  memberName: 'MemberName',
  status: 'Status',
  requestedAt: 'RequestedAt',
  notes: 'Notes'
} as const;

export const GOOGLE_CLIENT_ID_PROPERTY_KEY = 'GOOGLE_CLIENT_ID';
