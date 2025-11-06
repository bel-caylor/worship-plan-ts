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
  leader: 'Leader',
  preacher: 'Preacher',
  scripture: 'Scripture',
  scriptureText: 'Scripture Text',
  theme: 'Theme',
  keywords: 'Keywords',
  notes: 'Notes'
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
