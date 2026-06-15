/**
 * Helper utilities for Keeper Teams App
 */

/**
 * Check if a string looks like a Keeper UID (22 characters, base64-like)
 * @param {string} str - String to check
 * @returns {boolean}
 */
function isUid(str) {
  // Keeper UIDs are typically 22 characters, URL-safe base64
  const uidRegex = /^[A-Za-z0-9_-]{20,24}$/;
  return uidRegex.test(str);
}

/**
 * Check if string looks like a UID but has invalid length
 * @param {string} str - String to check
 * @returns {boolean}
 */
function looksLikeInvalidUid(str) {
  if (!str || typeof str !== 'string') return false;
  
  // If it's already a valid UID, return false
  if (isUid(str)) return false;

  const uidLikeRegex = /^[A-Za-z0-9_-]+$/;
  const hasNoSpaces = !str.includes(' ');
  const isBase64Like = uidLikeRegex.test(str);
  const length = str.length;

  return hasNoSpaces && isBase64Like && length >= 15 && (length < 20 || length > 24);
}

/**
 * Check if an error message indicates a permission conflict that requires manual revocation.
 * These errors mean the user already has access that conflicts with the new permission.
 * @param {string} errorMessage - Error message from Keeper
 * @returns {boolean}
 */
function isPermissionConflictError(errorMessage) {
  if (!errorMessage || typeof errorMessage !== 'string') return false;
  
  const errorLower = errorMessage.toLowerCase();
  
  const conflictIndicators = [
    'already has temporary access',
    'already has existing permissions',
    'conflicts with the selected permission',
    'conflicts with this permission',
    'first remove the user',
    'first revoke the user',
    'remove existing access',
    'revoke existing access',
    'time-limited access',
    'user already has',
  ];
  
  return conflictIndicators.some(indicator => errorLower.includes(indicator));
}

/**
 * Check if an error indicates the user is the record/folder owner.
 * @param {string} errorMessage - Error message from Keeper
 * @returns {boolean}
 */
function isRecordOwnerError(errorMessage) {
  if (!errorMessage || typeof errorMessage !== 'string') return false;
  
  const errorLower = errorMessage.toLowerCase();
  
  return errorLower.includes('cannot grant access to record owner') ||
         errorLower.includes('already owns this record') ||
         errorLower.includes('is the owner') ||
         errorLower.includes('record owner') ||
         errorLower.includes('cannot modify permissions for record owner');
}

/**
 * Check if an error indicates that one-time shares are not available for PAM records.
 * @param {string} errorMessage - Error message from Keeper
 * @returns {boolean}
 */
function isPamRecordError(errorMessage) {
  if (!errorMessage || typeof errorMessage !== 'string') return false;
  
  const errorLower = errorMessage.toLowerCase();
  
  return errorLower.includes('not available for pam records') ||
         errorLower.includes('pam records') ||
         errorLower.includes('pam record');
}

/**
 * Check if a record type is a PAM User record (eligible for rotate-on-expire).
 * @param {string} recordType - Record type string from Keeper
 * @returns {boolean}
 */
function isPamUserRecordType(recordType) {
  if (!recordType || typeof recordType !== 'string') return false;
  return recordType.toLowerCase().startsWith('pamuser');
}

/**
 * Check if rotation-not-configured error was returned by Commander.
 * @param {string} errorMessage - Error message from Keeper
 * @returns {boolean}
 */
function isRotationNotConfiguredError(errorMessage) {
  if (!errorMessage || typeof errorMessage !== 'string') return false;
  const lower = errorMessage.toLowerCase();
  return (
    lower.includes('rotation must be already set') ||
    (lower.includes('rotate') && lower.includes('expiration') && lower.includes('set on the record')) ||
    (lower.includes('--rotate-on-expiration') && (lower.includes('requires') || lower.includes('ineligible')))
  );
}

/**
 * Nested Share Folder (NSF) permission roles, in order of increasing privilege.
 * Shared by nsf-share-record and nsf-share-folder commands.
 */
const NSF_ROLES = ['viewer', 'share-manager', 'content-manager', 'content-share-manager', 'full-manager'];

/**
 * Check if a Keeper folder/record "type" string denotes a Nested Share Folder item.
 * Commander reports NSF items with a type containing "nested share".
 * @param {string} type - Type string from Keeper (e.g. 'nested share folder')
 * @returns {boolean}
 */
function isNsfType(type) {
  if (!type || typeof type !== 'string') return false;
  // Commander reports NSF folders in several string forms depending on the
  // command: `search` returns "nested_share_folder" (underscores), the
  // share-report Type column returns "Nested Share Folder" (spaces), and
  // subfolder listings embed "[Nested Share Folder]" in the name. Normalize
  // underscores to spaces so every variant is detected.
  return type.toLowerCase().replace(/_/g, ' ').includes('nested share');
}

/**
 * Validate an NSF permission role against the allowed set.
 * @param {string} role - Role string (e.g. 'viewer', 'full-manager')
 * @returns {boolean}
 */
function isValidNsfRole(role) {
  return typeof role === 'string' && NSF_ROLES.includes(role);
}

/**
 * Check whether a record's "Record Category" denotes a Nested Share Folder
 * record. Commander surfaces this in the search `details` string as e.g.
 * "Record Category: Nested". Unlike folders, NSF records keep their normal
 * record type (e.g. "login"), so the category is the reliable NSF signal.
 * @param {string} category - Category string from the record details
 * @returns {boolean}
 */
function isNsfRecordCategory(category) {
  if (!category || typeof category !== 'string') return false;
  const c = category.trim().toLowerCase();
  return c === 'nested' || c === 'keeperdrive';
}

/**
 * Convert a Date, epoch (ms), or date string into a strict RFC 3339 UTC string
 * (e.g. "2026-06-11T07:51:47Z") suitable for Adaptive Card DATE()/TIME() functions.
 * Legacy "YYYY-MM-DD HH:MM:SS" strings are interpreted as UTC.
 * @param {Date|number|string} input
 * @returns {string|null} RFC 3339 UTC string, or null if unparseable
 */
function toRfc3339Utc(input) {
  let date;
  if (input instanceof Date) {
    date = input;
  } else if (typeof input === 'number') {
    date = new Date(input);
  } else if (typeof input === 'string') {
    const trimmed = input.trim();
    if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(trimmed)) {
      // Legacy UTC timestamp without timezone marker.
      date = new Date(`${trimmed.replace(' ', 'T')}Z`);
    } else {
      date = new Date(trimmed);
    }
  } else {
    return null;
  }
  if (!date || isNaN(date.getTime())) return null;
  // Adaptive Cards expects no milliseconds (e.g. "...:47Z").
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/**
 * Format a timestamp for display in an Adaptive Card using Teams' native
 * DATE()/TIME() functions, so each viewer sees the value in THEIR OWN local
 * timezone. Server-side formatting cannot know the viewer's timezone, so this
 * is the only reliable way to localize times per recipient.
 *
 * Example: "{{DATE(2026-06-11T07:51:47Z, SHORT)}} {{TIME(2026-06-11T07:51:47Z)}}"
 *
 * @param {Date|number|string} input - Date/epoch/ISO string (assumed UTC if no tz)
 * @param {object} [options]
 * @param {string} [options.dateFormat='SHORT'] - COMPACT | SHORT | LONG
 * @param {boolean} [options.includeTime=true] - Whether to append the TIME() token
 * @returns {string} Adaptive Card token string, or the original input if unparseable
 */
function formatCardDateTime(input, options = {}) {
  const { dateFormat = 'SHORT', includeTime = true } = options;
  const iso = toRfc3339Utc(input);
  if (!iso) return typeof input === 'string' ? input : '';
  const datePart = `{{DATE(${iso}, ${dateFormat})}}`;
  return includeTime ? `${datePart} {{TIME(${iso})}}` : datePart;
}

/**
 * Sanitize text to prevent URL injection attacks.
 * Removes colons and forward slashes that could create clickable hyperlinks in Teams.
 * Used for UIDs and justification text displayed in cards.
 * @param {string} text - Text to sanitize
 * @returns {string} - Sanitized text
 */
function sanitizeHyperlinks(text) {
  if (!text || typeof text !== 'string') return text || '';
  
  // Remove colons and forward slashes to prevent URL injection
  // These characters can create clickable hyperlinks in Teams
  let sanitized = text.replace(/:/g, '');
  sanitized = sanitized.replace(/\//g, '');
  
  return sanitized;
}

module.exports = {
  isUid,
  looksLikeInvalidUid,
  isPermissionConflictError,
  isRecordOwnerError,
  isPamRecordError,
  isPamUserRecordType,
  isRotationNotConfiguredError,
  NSF_ROLES,
  isNsfType,
  isValidNsfRole,
  isNsfRecordCategory,
  toRfc3339Utc,
  formatCardDateTime,
  sanitizeHyperlinks,
};
