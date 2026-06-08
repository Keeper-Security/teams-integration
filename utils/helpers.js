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
  sanitizeHyperlinks,
};
