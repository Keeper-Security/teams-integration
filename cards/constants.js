/**
 * Constants for Adaptive Card builders
 * Shared permission and duration options
 */

/**
 * Permission options for records
 */
const RECORD_PERMISSIONS = [
  { title: 'View Only', value: 'view_only' },
  { title: 'Can Edit', value: 'can_edit' },
  { title: 'Can Share (Permanent)', value: 'can_share' },
  { title: 'Edit & Share (Permanent)', value: 'edit_and_share' },
  { title: 'Change Owner (Permanent)', value: 'change_owner' },
];

/**
 * Permission options for folders
 */
const FOLDER_PERMISSIONS = [
  { title: 'No Permissions', value: 'no_permissions' },
  { title: 'Manage Users (Permanent)', value: 'manage_users' },
  { title: 'Manage Records', value: 'manage_records' },
  { title: 'Manage All (Permanent)', value: 'manage_all' },
];

/**
 * Duration options for time-limited access
 */
const DURATION_OPTIONS = [
  { title: '5 minutes', value: '5m' },
  { title: '10 minutes', value: '10m' },
  { title: '30 minutes', value: '30m' },
  { title: '1 hour', value: '1h' },
  { title: '4 hours', value: '4h' },
  { title: '8 hours', value: '8h' },
  { title: '24 hours', value: '24h' },
  { title: '7 days', value: '7d' },
  { title: '30 days', value: '30d' },
  { title: 'Permanent', value: 'permanent' },
];

/**
 * Duration options excluding Permanent (used for PAM User targets where
 * rotate-on-expire is incompatible with permanent access)
 */
const DURATION_OPTIONS_NO_PERMANENT = DURATION_OPTIONS.filter(o => o.value !== 'permanent');

/**
 * Share duration options (subset for one-time shares)
 */
const SHARE_DURATION_OPTIONS = [
  { title: '5 minutes', value: '5m' },
  { title: '10 minutes', value: '10m' },
  { title: '30 minutes', value: '30m' },
  { title: '1 hour', value: '1h' },
  { title: '4 hours', value: '4h' },
  { title: '24 hours', value: '24h' },
  { title: '7 days', value: '7d' },
];

/**
 * Self-destruct duration options for auto-deleting records
 */
const SELF_DESTRUCT_DURATION_OPTIONS = [
  { title: '1 hour', value: '1h' },
  { title: '24 hours', value: '24h' },
  { title: '7 days', value: '7d' },
  { title: '30 days', value: '30d' },
  { title: '90 days', value: '90d' },
];

/**
 * Default duration for access grants and shares
 */
const DEFAULT_DURATION = '24h';

module.exports = {
  RECORD_PERMISSIONS,
  FOLDER_PERMISSIONS,
  DURATION_OPTIONS,
  DURATION_OPTIONS_NO_PERMANENT,
  SHARE_DURATION_OPTIONS,
  SELF_DESTRUCT_DURATION_OPTIONS,
  DEFAULT_DURATION,
};
