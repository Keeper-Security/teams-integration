/**
 * Approval Handler Helpers
 * 
 * Shared utility functions for approval handlers.
 */

const { getChannelService, getApprovalActivityId, removeApprovalActivityId, createLogger } = require('../../services');

const log = createLogger('ApprovalHelpers');

/**
 * Duration string to seconds mapping
 */
const DURATION_MAP = {
  '5m': 300,
  '10m': 600,
  '30m': 1800,
  '1h': 3600,
  '4h': 14400,
  '8h': 28800,
  '24h': 86400,
  '7d': 604800,
  '30d': 2592000,
  'permanent': null,
};

// Permissions that are always permanent (no expiration allowed)
const RECORD_PERMANENT_PERMISSIONS = ['can_share', 'edit_and_share', 'change_owner'];
const FOLDER_PERMANENT_PERMISSIONS = ['manage_users', 'manage_all'];

/**
 * Parse duration string to seconds
 * @param {string} duration - Duration string (e.g., '1h', '7d', 'permanent')
 * @returns {number|null} - Seconds or null for permanent
 */
function parseDuration(duration) {
  return DURATION_MAP[duration] ?? 86400;
}

/**
 * Get display duration based on permission type
 * For permanent-only permissions, always show 'Permanent' regardless of selected duration
 * @param {string} permission - The permission level
 * @param {string} duration - The selected duration string
 * @param {string} itemType - 'record' or 'folder'
 * @returns {string} - Display duration string
 */
function getDisplayDuration(permission, duration, itemType = 'record') {
  const permanentPermissions = itemType === 'folder' 
    ? FOLDER_PERMANENT_PERMISSIONS 
    : RECORD_PERMANENT_PERMISSIONS;
  
  if (permanentPermissions.includes(permission) || duration === 'permanent') {
    return 'Permanent';
  }
  return duration;
}

/**
 * Safely format expiry date, handling permanent/never strings
 * @param {string|null} expiresAt - The expiry date string from API
 * @returns {string} - Formatted date or 'Access granted indefinitely'
 */
function formatExpiryDate(expiresAt) {
  if (!expiresAt) return 'Access granted indefinitely';
  
  if (typeof expiresAt === 'string') {
    const lower = expiresAt.toLowerCase();
    if (lower.includes('permanent') || lower.includes('never') || lower === 'n/a') {
      return 'Access granted indefinitely';
    }
  }
  
  const expiryDate = new Date(expiresAt);
  if (isNaN(expiryDate.getTime())) {
    return expiresAt;
  }
  
  return expiryDate.toLocaleString('en-US', {
    month: '2-digit',
    day: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false,
  });
}

/**
 * Build notification card for requester when share invitation is sent
 * (user doesn't have Keeper account yet)
 */
function buildInvitationNotificationCard({ recordTitle, itemType, permission, approverName }) {
  const itemLabel = itemType === 'folder' ? 'folder' : 'record';
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      { 
        type: 'TextBlock', 
        text: 'Share Invitation Sent', 
        weight: 'Bolder', 
        size: 'Large',
        color: 'Warning'
      },
      {
        type: 'TextBlock',
        text: `Your request for **${itemLabel}** \`${recordTitle}\` has been approved!`,
        wrap: true,
        spacing: 'Medium'
      },
      {
        type: 'TextBlock',
        text: `However, you don't have a Keeper account yet. A share invitation has been sent to your email.`,
        wrap: true,
        size: 'Small'
      },
      {
        type: 'FactSet',
        spacing: 'Medium',
        facts: [
          { title: 'Status:', value: 'Invitation Sent' },
          { title: 'Permission:', value: permission || 'View Only' },
          { title: 'Approved by:', value: approverName },
        ],
      },
      {
        type: 'TextBlock',
        text: `**Next steps:** Check your email and accept the invitation to create your Keeper account and access the shared ${itemLabel}.`,
        wrap: true,
        spacing: 'Medium',
        size: 'Small',
        isSubtle: true
      },
    ],
    actions: [],
  };
}

/**
 * Build notification card for approver when permission conflict occurs
 * (user already has conflicting access that needs to be revoked first)
 */
function buildPermissionConflictCard({ itemTitle, itemType, requesterName, requesterEmail, requestedPermission, errorMessage }) {
  const itemLabel = itemType === 'folder' ? 'Folder' : 'Record';
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      { 
        type: 'TextBlock', 
        text: 'Permission Conflict', 
        weight: 'Bolder', 
        size: 'Large',
        color: 'Attention'
      },
      {
        type: 'TextBlock',
        text: `The approval could not be completed due to an existing permission conflict.`,
        wrap: true,
        spacing: 'Medium'
      },
      {
        type: 'FactSet',
        spacing: 'Medium',
        facts: [
          { title: `${itemLabel}:`, value: itemTitle || 'Unknown' },
          { title: 'Requester:', value: `${requesterName} (${requesterEmail})` },
          { title: 'Requested Permission:', value: requestedPermission || 'N/A' },
        ],
      },
      {
        type: 'Container',
        style: 'attention',
        spacing: 'Medium',
        items: [
          { type: 'TextBlock', text: 'Action Required:', weight: 'Bolder', size: 'Medium' },
          { 
            type: 'TextBlock', 
            text: `The user already has existing access to this ${itemLabel.toLowerCase()} that conflicts with the requested permission.`,
            size: 'Small',
            wrap: true 
          },
          { 
            type: 'TextBlock', 
            text: `To grant the new permission, you must first revoke their existing access using Keeper Commander or the Keeper web vault, then approve the request again.`,
            size: 'Small',
            wrap: true 
          },
        ],
      },
      {
        type: 'TextBlock',
        text: `**Error details:** ${errorMessage || 'Permission conflict detected'}`,
        wrap: true,
        spacing: 'Medium',
        size: 'Small',
        isSubtle: true
      },
    ],
    actions: [],
  };
}

/**
 * Get approver info from activity
 */
function getApproverInfo(activity) {
  const from = activity.from || {};
  return {
    name: from.name || 'Admin',
    id: from.id || 'unknown',
  };
}

/**
 * Helper function to update an approval card in the channel
 * Uses context.updateActivity() directly for reliable message updates
 * When the message is edited, Teams auto-refreshes the card for all users
 * @param {string} approvalId - The approval request ID
 * @param {Object} updatedCard - The updated Adaptive Card content
 * @param {Object} context - The Teams context
 */
async function tryUpdateApprovalCard(approvalId, updatedCard, context) {
  const activityId = context?.activity?.replyToId || getApprovalActivityId(approvalId);
  
  if (!activityId) {
    log.debug(`No activity ID found for approval ${approvalId}`);
    throw new Error('No activity ID found for this approval');
  }

  log.debug(`Updating approval ${approvalId}`, { activityId, hasContext: !!context });

  if (context && typeof context.updateActivity === 'function') {
    try {
      const updateActivity = {
        type: 'message',
        id: activityId,
        attachments: [{
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: updatedCard,
        }],
      };
      
      await context.updateActivity(updateActivity);
      removeApprovalActivityId(approvalId);
      log.info('Updated card via context.updateActivity()', { activityId });
      return true;
    } catch (contextError) {
      log.debug('context.updateActivity() not available or failed', contextError.message);
    }
  }

  const channelService = getChannelService();
  if (!channelService) {
    throw new Error('Channel service not available');
  }

  const success = await channelService.updateApprovalCard(activityId, updatedCard);
  if (success) {
    removeApprovalActivityId(approvalId);
    log.info('Updated card via channelService', { activityId });
    return true;
  }

  throw new Error('Failed to update channel card');
}

/**
 * Get current timestamp formatted for display
 * Returns format: "YYYY-MM-DD HH:MM:SS"
 * @returns {string} Formatted timestamp
 */
function getCurrentTimestamp() {
  return new Date().toISOString().replace('T', ' ').substring(0, 19);
}


function sanitizeDisplayField(val) {
  if (typeof val !== 'string') return val;
  return val.replace(/[<>&"']/g, '');
}

const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;


function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  return EMAIL_REGEX.test(email.trim());
}

module.exports = {
  DURATION_MAP,
  RECORD_PERMANENT_PERMISSIONS,
  FOLDER_PERMANENT_PERMISSIONS,
  parseDuration,
  getDisplayDuration,
  formatExpiryDate,
  buildInvitationNotificationCard,
  buildPermissionConflictCard,
  getApproverInfo,
  tryUpdateApprovalCard,
  getCurrentTimestamp,
  sanitizeDisplayField,
  isValidEmail,
};
