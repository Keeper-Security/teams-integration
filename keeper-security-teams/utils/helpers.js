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
 * Format a date for display
 * @param {Date|string} date - Date to format
 * @returns {string}
 */
function formatDate(date) {
  const d = typeof date === 'string' ? new Date(date) : date;
  return d.toLocaleString('en-US', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });
}

/**
 * Format duration in seconds to human-readable string
 * @param {number} seconds - Duration in seconds
 * @returns {string}
 */
function formatDuration(seconds) {
  if (seconds === null) return 'Permanent';
  
  if (seconds < 3600) {
    const minutes = Math.floor(seconds / 60);
    return `${minutes} minute${minutes !== 1 ? 's' : ''}`;
  }
  
  if (seconds < 86400) {
    const hours = Math.floor(seconds / 3600);
    return `${hours} hour${hours !== 1 ? 's' : ''}`;
  }
  
  const days = Math.floor(seconds / 86400);
  return `${days} day${days !== 1 ? 's' : ''}`;
}

/**
 * Parse command arguments, handling quoted strings
 * @param {string} text - Command text
 * @returns {string[]} - Array of arguments
 */
function parseArgs(text) {
  const args = [];
  let current = '';
  let inQuotes = false;
  let quoteChar = '';

  for (const char of text) {
    if ((char === '"' || char === "'") && !inQuotes) {
      inQuotes = true;
      quoteChar = char;
    } else if (char === quoteChar && inQuotes) {
      inQuotes = false;
      quoteChar = '';
    } else if (char === ' ' && !inQuotes) {
      if (current) {
        args.push(current);
        current = '';
      }
    } else {
      current += char;
    }
  }

  if (current) {
    args.push(current);
  }

  return args;
}

/**
 * Truncate a string to a maximum length with ellipsis
 * @param {string} str - String to truncate
 * @param {number} maxLength - Maximum length
 * @returns {string}
 */
function truncate(str, maxLength = 50) {
  if (str.length <= maxLength) return str;
  return str.substring(0, maxLength - 3) + '...';
}

/**
 * Get user display name from Teams activity
 * @param {Object} activity - Teams activity object
 * @returns {string}
 */
function getUserName(activity) {
  return activity?.from?.name || activity?.from?.id || 'Unknown User';
}

/**
 * Get user ID from Teams activity
 * @param {Object} activity - Teams activity object
 * @returns {string}
 */
function getUserId(activity) {
  return activity?.from?.id || '';
}

/**
 * Sleep for a specified duration
 * @param {number} ms - Milliseconds to sleep
 * @returns {Promise<void>}
 */
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Retry a function with exponential backoff
 * @param {Function} fn - Function to retry
 * @param {number} maxRetries - Maximum number of retries
 * @param {number} baseDelay - Base delay in milliseconds
 * @returns {Promise<*>}
 */
async function retry(fn, maxRetries = 3, baseDelay = 1000) {
  let lastError;
  
  for (let i = 0; i < maxRetries; i++) {
    try {
      return await fn();
    } catch (error) {
      lastError = error;
      if (i < maxRetries - 1) {
        const delay = baseDelay * Math.pow(2, i);
        await sleep(delay);
      }
    }
  }
  
  throw lastError;
}

/**
 * Send a direct message to a user
 * Uses the channel service to send proactive messages
 * @param {string} userId - Teams user ID
 * @param {string|Object} message - Text message or adaptive card
 * @returns {Promise<boolean>} - Success status
 */
async function sendDirectMessage(userId, message) {
  const { getChannelService } = require('../services');
  const channelService = getChannelService();
  
  if (!channelService) {
    console.warn('[Helper] Channel service not initialized');
    return false;
  }
  
  return await channelService.sendDirectMessage(userId, message);
}

/**
 * Send access granted notification to user
 * @param {string} userId - Teams user ID
 * @param {Object} options - Notification options
 * @returns {Promise<boolean>}
 */
async function sendAccessGrantedNotification(userId, {
  approvalId,
  itemType,
  itemTitle,
  uid,
  permission,
  expiresAt,
  approverName,
}) {
  const message = 
    `✅ **Access Granted!**\n\n` +
    `Your request has been approved.\n\n` +
    `• **Request ID:** \`${approvalId}\`\n` +
    `• **${itemType}:** ${itemTitle}\n` +
    `• **UID:** \`${uid}\`\n` +
    `• **Permission:** ${permission}\n` +
    `• **Expires:** ${expiresAt}\n` +
    `• **Approved by:** ${approverName}\n\n` +
    `Check your Keeper vault for access.`;
  
  return await sendDirectMessage(userId, message);
}

/**
 * Send access denied notification to user
 * @param {string} userId - Teams user ID
 * @param {Object} options - Notification options
 * @returns {Promise<boolean>}
 */
async function sendAccessDeniedNotification(userId, {
  approvalId,
  itemType,
  itemTitle,
  denierName,
}) {
  const message = 
    `❌ **Access Request Denied**\n\n` +
    `Your request was not approved.\n\n` +
    `• **Request ID:** \`${approvalId}\`\n` +
    `• **${itemType}:** ${itemTitle}\n` +
    `• **Denied by:** ${denierName}\n\n` +
    `If you believe this was in error, please contact your administrator.`;
  
  return await sendDirectMessage(userId, message);
}

/**
 * Send one-time share link to user
 * @param {string} userId - Teams user ID
 * @param {Object} options - Share options
 * @returns {Promise<boolean>}
 */
async function sendShareLinkNotification(userId, {
  approvalId,
  recordTitle,
  recordUid,
  shareUrl,
  expiresAt,
  approverName,
}) {
  const message = 
    `**One-Time Share Link Created**\n\n` +
    `Your share request has been approved!\n\n` +
    `• **Request ID:** \`${approvalId}\`\n` +
    `• **Record:** ${recordTitle}\n` +
    `• **UID:** \`${recordUid}\`\n` +
    `• **Approved by:** ${approverName}\n\n` +
    `**Share Link:**\n${shareUrl}\n\n` +
    `**Expires:** ${expiresAt}\n\n` +
    `**Security Notice:**\n` +
    `• This link can only be opened on ONE device\n` +
    `• It expires after first access or time limit\n` +
    `• Share only via secure channels\n` +
    `• Do NOT post in public channels`;
  
  return await sendDirectMessage(userId, message);
}

/**
 * Generate a unique approval ID
 * @returns {string}
 */
function generateApprovalId() {
  const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const randomStr = Math.random().toString(36).substring(2, 7).toUpperCase();
  return `APR-${dateStr}-${randomStr}`;
}

/**
 * Parse duration string to seconds
 * @param {string} duration - Duration string (e.g., '1h', '24h', '7d', 'permanent')
 * @returns {number|null} - Seconds or null for permanent
 */
function parseDurationToSeconds(duration) {
  if (duration === 'permanent' || !duration) return null;
  
  const mapping = {
    '1h': 3600,
    '4h': 14400,
    '8h': 28800,
    '24h': 86400,
    '7d': 604800,
    '30d': 2592000,
  };
  
  return mapping[duration] || 3600; // Default 1 hour
}

/**
 * Format permission level for display
 * @param {string} permission - Permission value
 * @returns {string}
 */
function formatPermission(permission) {
  const mappings = {
    'view_only': 'View Only',
    'can_edit': 'Can Edit',
    'can_share': 'Can Share',
    'edit_and_share': 'Edit & Share',
    'change_owner': 'Change Owner',
    'no_permissions': 'No Permissions',
    'manage_users': 'Manage Users',
    'manage_records': 'Manage Records',
    'manage_all': 'Manage All',
  };
  
  return mappings[permission] || permission;
}

module.exports = {
  isUid,
  formatDate,
  formatDuration,
  parseArgs,
  truncate,
  getUserName,
  getUserId,
  sleep,
  retry,
  sendDirectMessage,
  sendAccessGrantedNotification,
  sendAccessDeniedNotification,
  sendShareLinkNotification,
  generateApprovalId,
  parseDurationToSeconds,
  formatPermission,
};
