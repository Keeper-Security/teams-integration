/**
 * Type definitions and constants for Keeper Teams App
 * 
 * Note: These are JSDoc types for documentation and IDE support.
 * In a TypeScript project, these would be actual interfaces.
 */

/**
 * Request types for access requests
 * @enum {string}
 */
const RequestType = {
  RECORD: 'record',
  FOLDER: 'folder',
  ONE_TIME_SHARE: 'one_time_share',
};

/**
 * Permission levels for record access
 * @enum {string}
 */
const RecordPermission = {
  VIEW_ONLY: 'view_only',
  CAN_EDIT: 'can_edit',
  CAN_SHARE: 'can_share',
  EDIT_AND_SHARE: 'edit_and_share',
  CHANGE_OWNER: 'change_owner',
};

/**
 * Permission levels for folder access
 * @enum {string}
 */
const FolderPermission = {
  NO_PERMISSIONS: 'no_permissions',
  MANAGE_USERS: 'manage_users',
  MANAGE_RECORDS: 'manage_records',
  MANAGE_ALL: 'manage_all',
};

/**
 * Permissions that require permanent access (no time limit)
 */
const PERMANENT_ONLY_PERMISSIONS = [
  RecordPermission.CAN_SHARE,
  RecordPermission.EDIT_AND_SHARE,
  RecordPermission.CHANGE_OWNER,
  FolderPermission.MANAGE_USERS,
  FolderPermission.MANAGE_ALL,
];

/**
 * Duration options with their values in seconds
 */
const DURATION_SECONDS = {
  '1h': 3600,
  '4h': 14400,
  '8h': 28800,
  '24h': 86400,
  '7d': 604800,
  '30d': 2592000,
  'permanent': null,
};

module.exports = {
  RequestType,
  RecordPermission,
  FolderPermission,
  PERMANENT_ONLY_PERMISSIONS,
  DURATION_SECONDS,
};
