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

module.exports = {
  isUid,
};
