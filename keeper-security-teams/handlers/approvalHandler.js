/**
 * Approval Action Handler
 * 
 * Handles Adaptive Card action submissions for:
 * - Approve/Deny record access requests
 * - Approve/Deny folder access requests
 * - Approve/Deny one-time share requests
 */

const keeperClient = require('../services/keeperClient');
const cards = require('../cards');
const config = require('../config');
const { getChannelService, getApprovalActivityId, removeApprovalActivityId, getApprovalStatus, storeApprovalStatus, createLogger } = require('../services');

const log = createLogger('ApprovalHandler');

/**
 * Duration string to seconds mapping
 */
const DURATION_MAP = {
  '1h': 3600,
  '4h': 14400,
  '8h': 28800,
  '24h': 86400,
  '7d': 604800,
  '30d': 2592000,
  'permanent': null,
};

/**
 * Parse duration string to seconds
 * @param {string} duration - Duration string (e.g., '1h', '7d', 'permanent')
 * @returns {number|null} - Seconds or null for permanent
 */
function parseDuration(duration) {
  return DURATION_MAP[duration] ?? 86400; // Default to 24h
}

/**
 * Safely format expiry date, handling permanent/never strings
 * @param {string|null} expiresAt - The expiry date string from API
 * @returns {string} - Formatted date or 'Access granted indefinitely'
 */
function formatExpiryDate(expiresAt) {
  if (!expiresAt) return 'Access granted indefinitely';
  
  // Check for permanent/never strings
  if (typeof expiresAt === 'string') {
    const lower = expiresAt.toLowerCase();
    if (lower.includes('permanent') || lower.includes('never') || lower === 'n/a') {
      return 'Access granted indefinitely';
    }
  }
  
  // Try to parse as date
  const expiryDate = new Date(expiresAt);
  if (isNaN(expiryDate.getTime())) {
    return expiresAt; // Return original if not a valid date
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
    version: '1.4',
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
        spacing: 'Small'
      },
      {
        type: 'Container',
        style: 'emphasis',
        spacing: 'Medium',
        items: [
          { type: 'TextBlock', text: 'Next Steps:', weight: 'Bolder', size: 'Medium' },
          { type: 'TextBlock', text: '1. Check your email for the Keeper invitation', size: 'Small', wrap: true },
          { type: 'TextBlock', text: '2. Accept the invitation and create a Keeper account', size: 'Small', wrap: true },
          { type: 'TextBlock', text: `3. The ${itemLabel} will be automatically shared with you`, size: 'Small', wrap: true },
        ],
      },
      {
        type: 'FactSet',
        spacing: 'Medium',
        facts: [
          { title: 'Permission:', value: permission || 'View Only' },
          { title: 'Approved by:', value: approverName || 'Admin' },
        ],
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
 * Helper function to update an approval card using the context's activity info
 * This uses the context directly since it has the right service URL and permissions
 * @param {string} approvalId - The approval request ID
 * @param {Object} updatedCard - The updated Adaptive Card content
 * @param {Object} context - The Teams context (required)
 */
async function tryUpdateApprovalCard(approvalId, updatedCard, context) {
  if (!context || !context.activity) {
    throw new Error('Context with activity is required for updates');
  }

  const activity = context.activity;
  
  // Get activity ID - try stored first, then replyToId
  let activityId = getApprovalActivityId(approvalId);
  
  if (!activityId && activity.replyToId) {
    log.debug(`Using replyToId as fallback: ${activity.replyToId}`);
    activityId = activity.replyToId;
  }
  
  if (!activityId) {
    log.debug(`No activity ID found for approval ${approvalId}`);
    throw new Error('No activity ID found for this approval');
  }

  // Get conversation ID - remove any messageid suffix
  let conversationId = activity.conversation?.id;
  if (conversationId && conversationId.includes(';messageid=')) {
    conversationId = conversationId.split(';messageid=')[0];
  }

  // Get service URL from the context's activity
  const serviceUrl = activity.serviceUrl;

  if (!conversationId || !serviceUrl) {
    throw new Error('Missing conversation ID or service URL from context');
  }

  log.debug(`Updating approval ${approvalId}`, { activityId });

  // Build the updated activity
  const updatedActivity = {
    type: 'message',
    id: activityId,
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: updatedCard,
    }],
  };

  // Try multiple methods to update the activity
  
  // Method 1: Try using context's send with the updated activity (replace behavior)
  // Some SDKs support this pattern
  if (typeof context.updateActivity === 'function') {
    try {
      await context.updateActivity(updatedActivity);
      log.debug('Updated via context.updateActivity');
      removeApprovalActivityId(approvalId);
      return true;
    } catch (e) {
      log.debug('context.updateActivity failed', e.message);
    }
  }

  // Method 2: Try using the channel service with context info
  const channelService = getChannelService();
  if (channelService && channelService.app) {
    try {
      // Create API client with the context's service URL
      const { Client: ApiClient } = require('@microsoft/teams.api');
      const apiClient = new ApiClient(serviceUrl, channelService.app.client);
      
      const result = await apiClient.conversations.activities(conversationId).update(
        activityId,
        updatedActivity
      );
      
      log.debug('Updated via ApiClient with context serviceUrl');
      removeApprovalActivityId(approvalId);
      return true;
    } catch (e) {
      log.debug('ApiClient update failed', e.message);
    }
  }

  // Method 3: Try using the channel service's updateApprovalCard as last resort
  if (channelService) {
    try {
      const success = await channelService.updateApprovalCard(activityId, updatedCard);
      if (success) {
        removeApprovalActivityId(approvalId);
        log.debug('Updated via channelService');
        return true;
      }
    } catch (e) {
      log.debug('channelService.updateApprovalCard failed', e.message);
    }
  }

  throw new Error('All update methods failed');
}

/**
 * Handle approval of a record access request
 */
async function handleRecordApproval(context, data) {
  log.debug('handleRecordApproval called', data);
  
  const approver = getApproverInfo(context.activity);
  const permission = data.permission || 'view_only';
  const duration = data.duration || '24h';
  const durationSeconds = parseDuration(duration);
  
  const recordUid = data.recordUid;
  const recordTitle = data.recordTitle || recordUid;
  const requesterEmail = data.requesterEmail;
  const requesterName = data.requesterName || 'User';
  const approvalId = data.approvalId;
  const justification = data.justification;
  
  if (!recordUid) {
    await context.send('Error: Missing record UID');
    return;
  }
  
  if (!requesterEmail) {
    await context.send('Error: Missing requester email. Cannot grant access without email.');
    return;
  }
  
  // Grant access
  const result = await keeperClient.grantRecordAccess(
    recordUid,
    requesterEmail,
    permission,
    durationSeconds
  );
  
  if (result.success) {
    // Format expiry date using helper function
    const expiresAtFormatted = formatExpiryDate(result.expiresAt);
    const processedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
    
    let updatedCard;
    
    if (result.invitationSent) {
      log.info('Share invitation sent to user (no Keeper account)');
      
      // Build invitation sent card
      updatedCard = cards.buildRecordInvitationSentCard({
        approvalId: approvalId,
        requesterName: requesterName,
        requesterEmail: requesterEmail,
        recordTitle: recordTitle,
        recordUid: recordUid,
        justification: justification,
        permission: permission,
        approverName: approver.name,
        processedTime: processedTime,
      });
    } else {
      // Build the updated card with APPROVED status and all details
      updatedCard = cards.buildRecordApprovalCardWithStatus({
        approvalId: approvalId,
        requesterName: requesterName,
        requesterEmail: requesterEmail,
        recordTitle: recordTitle,
        justification: justification,
        status: 'approved',
        approverName: approver.name,
        permission: permission,
        duration: duration === 'permanent' ? 'Permanent' : duration,
        expiresAt: expiresAtFormatted,
      });
    }
    
    try {
      await tryUpdateApprovalCard(approvalId, updatedCard, context);
      log.debug('Successfully updated approval card with', result.invitationSent ? 'INVITATION SENT' : 'APPROVED', 'status');
    } catch (updateError) {
      log.debug('Failed to update activity, sending new message', updateError.message);
      // Fallback: send a new message if update fails
      await context.send({
        type: 'message',
        attachments: [{
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: updatedCard,
        }],
      });
    }
    
    // Send DM notification to the requester
    const requesterId = data.requesterId;
    if (requesterId) {
      try {
        const channelService = getChannelService();
        if (channelService) {
          let notificationCard;
          
          if (result.invitationSent) {
            // Send invitation notification to requester
            notificationCard = buildInvitationNotificationCard({
              recordTitle: recordTitle,
              itemType: 'record',
              permission: permission,
              approverName: approver.name,
            });
          } else {
            // Send regular approval notification
            notificationCard = cards.buildRequesterNotificationCard({
              approved: true,
              recordTitle: recordTitle,
              permission: permission,
              duration: duration === 'permanent' ? 'Permanent' : duration,
              expiresAt: expiresAtFormatted,
              approverName: approver.name,
            });
          }
          
          const notificationSent = await channelService.sendDirectMessage(requesterId, {
            type: 'message',
            attachments: [{
              contentType: 'application/vnd.microsoft.card.adaptive',
              content: notificationCard,
            }],
          });
          
          if (notificationSent) {
            log.debug(`Sent ${result.invitationSent ? 'invitation' : 'approval'} notification to requester: ${requesterId}`);
          } else {
            log.debug(`Could not send notification to requester (no reference stored)`);
          }
        }
      } catch (notifyError) {
        log.error('Error sending requester notification', notifyError.message);
      }
    }
  } else {
    await context.send('Failed to grant access: ' + result.error);
  }
}

/**
 * Handle denial of a record access request
 */
async function handleRecordDenial(context, data) {
  log.debug('handleRecordDenial called', data);
  
  const approver = getApproverInfo(context.activity);
  const recordTitle = data.recordTitle || data.recordUid || 'Unknown Record';
  const requesterName = data.requesterName || 'User';
  const approvalId = data.approvalId || 'N/A';
  const justification = data.justification || '';
  
  log.debug('Denying record access', { approver: approver.name, recordTitle, requesterName });
  
  // Build updated card with DENIED status
  const updatedCard = cards.buildRecordApprovalCardWithStatus({
    approvalId,
    requesterName,
    recordTitle,
    justification,
    status: 'denied',
    approverName: approver.name,
  });
  
  try {
    await tryUpdateApprovalCard(approvalId, updatedCard, context);
    log.debug('Updated original card with denied status');
  } catch (error) {
    log.debug('Failed to update card, sending new message', error.message);
    // Fallback: send as new message
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: updatedCard,
      }],
    });
  }
  
  // Send DM notification to the requester
  const requesterId = data.requesterId;
  if (requesterId) {
    try {
      const channelService = getChannelService();
      if (channelService) {
        const notificationCard = cards.buildRequesterNotificationCard({
          approved: false,
          recordTitle: recordTitle,
          approverName: approver.name,
          denialReason: data.denialReason || null,
        });
        
        const notificationSent = await channelService.sendDirectMessage(requesterId, {
          type: 'message',
          attachments: [{
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: notificationCard,
          }],
        });
        
        if (notificationSent) {
          log.debug(`Sent denial notification to requester: ${requesterId}`);
        } else {
          log.debug(`Could not send notification to requester (no reference stored)`);
        }
      }
    } catch (notifyError) {
      log.error('Error sending requester notification', notifyError.message);
    }
  }
  
  log.debug('Denial complete');
}

/**
 * Handle approval of a folder access request
 */
async function handleFolderApproval(context, data) {
  log.debug('handleFolderApproval called', data);
  
  const approver = getApproverInfo(context.activity);
  const permission = data.permission || 'no_permissions';
  const duration = data.duration || '24h';
  const durationSeconds = parseDuration(duration);
  
  const folderUid = data.folderUid;
  const folderName = data.folderName || folderUid;
  const requesterEmail = data.requesterEmail;
  const requesterName = data.requesterName || 'User';
  const approvalId = data.approvalId;
  const justification = data.justification;
  
  if (!folderUid) {
    await context.send('Error: Missing folder UID');
    return;
  }
  
  if (!requesterEmail) {
    await context.send('Error: Missing requester email. Cannot grant access without email.');
    return;
  }
  
  // Grant access
  const result = await keeperClient.grantFolderAccess(
    folderUid,
    requesterEmail,
    permission,
    durationSeconds
  );
  
  if (result.success) {
    // Format expiry date using helper function
    const expiresAtFormatted = formatExpiryDate(result.expiresAt);
    const processedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
    
    let updatedCard;
    
    if (result.invitationSent) {
      log.info('Share invitation sent to user for folder (no Keeper account)');
      
      // Build invitation sent card
      updatedCard = cards.buildFolderInvitationSentCard({
        approvalId: approvalId,
        requesterName: requesterName,
        requesterEmail: requesterEmail,
        folderName: folderName,
        folderUid: folderUid,
        justification: justification,
        permission: permission,
        approverName: approver.name,
        processedTime: processedTime,
      });
    } else {
      // Build the updated card with APPROVED status
      updatedCard = cards.buildFolderApprovalCardWithStatus({
        approvalId: approvalId,
        requesterName: requesterName,
        requesterEmail: requesterEmail,
        folderName: folderName,
        justification: justification,
        status: 'approved',
        approverName: approver.name,
        permission: permission,
        duration: duration === 'permanent' ? 'Permanent' : duration,
        expiresAt: expiresAtFormatted,
      });
    }
    
    try {
      await tryUpdateApprovalCard(approvalId, updatedCard, context);
      log.debug(`Successfully updated folder approval card with ${result.invitationSent ? 'INVITATION SENT' : 'APPROVED'} status`);
    } catch (updateError) {
      log.debug('Failed to update activity, sending new message', updateError.message);
      // Fallback: send a new message if update fails
      await context.send({
        type: 'message',
        attachments: [{
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: updatedCard,
        }],
      });
    }
    
    // Send DM notification to the requester
    const requesterId = data.requesterId;
    if (requesterId) {
      try {
        const channelService = getChannelService();
        if (channelService) {
          let notificationCard;
          
          if (result.invitationSent) {
            // Send invitation notification to requester
            notificationCard = buildInvitationNotificationCard({
              recordTitle: folderName,
              itemType: 'folder',
              permission: permission,
              approverName: approver.name,
            });
          } else {
            // Send regular approval notification
            notificationCard = cards.buildRequesterNotificationCard({
              approved: true,
              recordTitle: folderName,
              itemType: 'folder',
              permission: permission,
              duration: duration === 'permanent' ? 'Permanent' : duration,
              expiresAt: expiresAtFormatted,
              approverName: approver.name,
            });
          }
          
          const notificationSent = await channelService.sendDirectMessage(requesterId, {
            type: 'message',
            attachments: [{
              contentType: 'application/vnd.microsoft.card.adaptive',
              content: notificationCard,
            }],
          });
          
          if (notificationSent) {
            log.debug(`Sent ${result.invitationSent ? 'invitation' : 'folder approval'} notification to requester: ${requesterId}`);
          } else {
            log.debug(`Could not send notification to requester (no reference stored)`);
          }
        }
      } catch (notifyError) {
        log.error('Error sending requester notification', notifyError.message);
      }
    }
  } else {
    await context.send('Failed to grant folder access: ' + result.error);
  }
}

/**
 * Handle denial of a folder access request
 */
async function handleFolderDenial(context, data) {
  log.debug('handleFolderDenial called', data);
  
  const approver = getApproverInfo(context.activity);
  const folderName = data.folderName || data.folderUid || 'Unknown Folder';
  const requesterName = data.requesterName || 'User';
  const approvalId = data.approvalId || 'N/A';
  const justification = data.justification || '';
  
  log.debug('Denying folder access', { approver: approver.name, folderName, requesterName });
  
  // Build updated card with DENIED status
  const updatedCard = cards.buildFolderApprovalCardWithStatus({
    approvalId,
    requesterName,
    folderName,
    justification,
    status: 'denied',
    approverName: approver.name,
  });
  
  try {
    await tryUpdateApprovalCard(approvalId, updatedCard, context);
    log.debug('Updated original folder card with denied status');
  } catch (error) {
    log.debug('Failed to update folder card, sending new message', error.message);
    // Fallback: send as new message
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: updatedCard,
      }],
    });
  }
  
  // Send DM notification to the requester
  const requesterId = data.requesterId;
  if (requesterId) {
    try {
      const channelService = getChannelService();
      if (channelService) {
        const notificationCard = cards.buildRequesterNotificationCard({
          approved: false,
          recordTitle: folderName,
          itemType: 'folder',
          approverName: approver.name,
          denialReason: data.denialReason || null,
        });
        
        const notificationSent = await channelService.sendDirectMessage(requesterId, {
          type: 'message',
          attachments: [{
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: notificationCard,
          }],
        });
        
        if (notificationSent) {
          log.debug(`Sent folder denial notification to requester: ${requesterId}`);
        } else {
          log.debug(`Could not send notification to requester (no reference stored)`);
        }
      }
    } catch (notifyError) {
      log.error('Error sending requester notification', notifyError.message);
    }
  }
  
  log.debug('Folder denial complete');
}

/**
 * Handle approval of a one-time share request
 */
async function handleShareApproval(context, data) {
  const approver = getApproverInfo(context.activity);
  const duration = data.duration || '24h';
  const durationSeconds = parseDuration(duration) || 86400;
  const editable = data.editable === 'true' || data.editable === true;
  
  const recordUid = data.recordUid;
  const recordTitle = data.recordTitle || recordUid;
  const requesterName = data.requesterName || 'User';
  
  if (!recordUid) {
    await context.send('Error: Missing record UID');
    return;
  }
  
  // Create the share
  await context.send('Creating one-time share link...');
  
  const result = await keeperClient.createOneTimeShare(
    recordUid,
    durationSeconds,
    editable
  );
  
  if (result.success) {
    const shareCard = cards.buildShareResultCard({
      success: true,
      recordTitle: recordTitle,
      shareUrl: result.shareUrl,
      expiresAt: result.expiresAt,
    });
    
    await context.send({
      type: 'message',
      text: 'Share link created for ' + requesterName,
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: shareCard,
      }],
    });
  } else {
    await context.send('Failed to create share link: ' + result.error);
  }
}

/**
 * Handle denial of a one-time share request
 */
async function handleShareDenial(context, data) {
  const approver = getApproverInfo(context.activity);
  const recordTitle = data.recordTitle || data.recordUid;
  const requesterName = data.requesterName || 'User';
  
  const deniedCard = cards.buildDeniedMessageCard(
    approver.name,
    null,
    recordTitle,
    'record'
  );
  
  await context.send({
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: deniedCard,
    }],
  });
  
  await context.send('Share request denied for ' + requesterName + ' for record **' + recordTitle + '**');
}

/**
 * Route card action to appropriate handler
 */
async function routeApprovalAction(context, data) {
  const action = data.action;
  
  switch (action) {
    case 'approve_record':
      await handleRecordApproval(context, data);
      return true;
      
    case 'deny_record':
      await handleRecordDenial(context, data);
      return true;
      
    case 'approve_folder':
      await handleFolderApproval(context, data);
      return true;
      
    case 'deny_folder':
      await handleFolderDenial(context, data);
      return true;
      
    case 'approve_share':
      await handleShareApproval(context, data);
      return true;
      
    case 'deny_share':
      await handleShareDenial(context, data);
      return true;
      
    default:
      return false;
  }
}

/**
 * Route card action and return updated card for Universal Actions
 * This is used with Action.Execute to return the updated card directly
 */
async function routeApprovalActionWithCardResponse(context, data) {
  const action = data.action;
  const approver = getApproverInfo(context.activity);
  const approvalId = data.approvalId;
  
  log.debug('routeApprovalActionWithCardResponse', action);
  
  // Check if this approval has already been processed
  if (approvalId) {
    const existingStatus = getApprovalStatus(approvalId);
    if (existingStatus) {
      log.debug(`Approval ${approvalId} already processed: ${existingStatus.status}`);
      
      const statusText = existingStatus.status === 'approved' ? 'APPROVED' : 'DENIED';
      const itemName = existingStatus.recordTitle || existingStatus.folderName || 'the requested item';
      const itemType = existingStatus.type === 'folder' ? 'Folder' : 'Record';
      
      // Return a card showing the approval was already processed
      const alreadyProcessedCard = {
        type: 'AdaptiveCard',
        '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
        version: '1.2',
        body: [
          {
            type: 'TextBlock',
            text: `This request has already been ${statusText.toLowerCase()}`,
            weight: 'Bolder',
            size: 'Large',
            wrap: true,
            color: existingStatus.status === 'approved' ? 'Good' : 'Attention',
          },
          {
            type: 'FactSet',
            facts: [
              { title: 'Status:', value: statusText },
              { title: `${itemType}:`, value: itemName },
              { title: 'Processed By:', value: existingStatus.approverName || 'Unknown' },
              { title: 'Time:', value: existingStatus.processedTime || existingStatus.updatedAt || 'Unknown' },
            ],
          },
          {
            type: 'TextBlock',
            text: 'No further action is needed.',
            wrap: true,
            isSubtle: true,
            spacing: 'Medium',
          },
        ],
        actions: [],
      };
      
      return { updatedCard: alreadyProcessedCard };
    }
  }
  
  switch (action) {
    case 'approve_record': {
      const { approvalId, recordUid, recordTitle, requesterName, requesterId, requesterEmail } = data;
      const permission = data.permission || 'view_only';
      const duration = data.duration || '24h';
      const durationSeconds = parseDuration(duration);
      const processedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
      
      log.debug('Approving record via Universal Action', { approver: approver.name, recordTitle, requesterName, permission, duration });
      
      if (!recordUid) {
        log.error('Missing record UID');
        return { error: 'Missing record UID' };
      }
      
      if (!requesterEmail) {
        log.error('Missing requester email');
        return { error: 'Missing requester email' };
      }
      
      // Grant access
      const result = await keeperClient.grantRecordAccess(
        recordUid,
        requesterEmail,
        permission,
        durationSeconds
      );
      
      if (!result.success) {
        log.error('Failed to grant access', result.error);
        return { error: result.error };
      }
      
      // Format expiry date using helper function
      const expiresAtFormatted = formatExpiryDate(result.expiresAt);
      
      // Check if invitation was sent (user doesn't have Keeper account yet)
      const isInvitationSent = result.invitationSent;
      
      // Store approval status for refresh mechanism
      storeApprovalStatus(approvalId, {
        status: isInvitationSent ? 'invitation_sent' : 'approved',
        type: 'record',
        approverName: approver.name,
        requesterName,
        requesterEmail,
        recordTitle,
        justification: data.justification || '',
        permission,
        duration: duration === 'permanent' ? 'Permanent' : duration,
        expiresAt: expiresAtFormatted,
        processedTime,
        invitationSent: isInvitationSent,
      });
      
      let updatedCard;
      
      if (isInvitationSent) {
        log.info('Share invitation sent for record (user has no Keeper account)');
        // Build invitation sent card
        updatedCard = cards.buildRecordInvitationSentCard({
          approvalId,
          requesterName,
          requesterEmail,
          recordTitle,
          recordUid,
          justification: data.justification || '',
          permission: permission,
          approverName: approver.name,
          processedTime,
        });
      } else {
        // Build updated card with APPROVED status
        updatedCard = cards.buildRecordApprovalCardWithStatus({
          approvalId,
          requesterName,
          requesterEmail,
          recordTitle,
          justification: data.justification || '',
          status: 'approved',
          approverName: approver.name,
          permission: permission,
          duration: duration === 'permanent' ? 'Permanent' : duration,
          expiresAt: expiresAtFormatted,
          processedTime,
        });
      }
      
      // Send notification to requester (async, don't block)
      if (requesterId) {
        try {
          const channelService = getChannelService();
          if (channelService) {
            let notificationCard;
            
            if (isInvitationSent) {
              notificationCard = buildInvitationNotificationCard({
                recordTitle: recordTitle,
                itemType: 'record',
                permission: permission,
                approverName: approver.name,
              });
            } else {
              notificationCard = cards.buildRequesterNotificationCard({
                approved: true,
                recordTitle: recordTitle,
                permission: permission,
                duration: duration === 'permanent' ? 'Permanent' : duration,
                expiresAt: expiresAtFormatted,
                approverName: approver.name,
                itemType: 'record',
              });
            }
            
            channelService.sendDirectMessage(requesterId, {
              type: 'message',
              attachments: [{
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: notificationCard,
              }],
            }).then(sent => {
              if (sent) {
                log.debug(`Sent ${isInvitationSent ? 'invitation' : 'approval'} notification to requester`);
              }
            }).catch(err => {
              log.debug(`Could not send notification`, err.message);
            });
          }
        } catch (notifyError) {
          log.error('Error sending notification', notifyError.message);
        }
      }
      
      return { updatedCard };
    }
    
    case 'approve_folder': {
      const { approvalId, folderUid, folderName, requesterName, requesterId, requesterEmail } = data;
      const permission = data.permission || 'no_permissions';
      const duration = data.duration || '24h';
      const durationSeconds = parseDuration(duration);
      const processedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
      
      log.debug('Approving folder via Universal Action', { approver: approver.name, folderName, requesterName, permission, duration });
      
      if (!folderUid) {
        log.error('Missing folder UID');
        return { error: 'Missing folder UID' };
      }
      
      if (!requesterEmail) {
        log.error('Missing requester email');
        return { error: 'Missing requester email' };
      }
      
      // Grant access
      const result = await keeperClient.grantFolderAccess(
        folderUid,
        requesterEmail,
        permission,
        durationSeconds
      );
      
      if (!result.success) {
        log.error('Failed to grant folder access', result.error);
        return { error: result.error };
      }
      
      // Format expiry date using helper function
      const expiresAtFormatted = formatExpiryDate(result.expiresAt);
      
      // Check if invitation was sent (user doesn't have Keeper account yet)
      const isInvitationSent = result.invitationSent;
      
      // Store approval status for refresh mechanism
      storeApprovalStatus(approvalId, {
        status: isInvitationSent ? 'invitation_sent' : 'approved',
        type: 'folder',
        approverName: approver.name,
        requesterName,
        requesterEmail,
        folderName,
        justification: data.justification || '',
        permission,
        duration: duration === 'permanent' ? 'Permanent' : duration,
        expiresAt: expiresAtFormatted,
        processedTime,
        invitationSent: isInvitationSent,
      });
      
      let updatedCard;
      
      if (isInvitationSent) {
        log.info('Share invitation sent for folder (user has no Keeper account)');
        // Build invitation sent card
        updatedCard = cards.buildFolderInvitationSentCard({
          approvalId,
          requesterName,
          requesterEmail,
          folderName,
          folderUid,
          justification: data.justification || '',
          permission: permission,
          approverName: approver.name,
          processedTime,
        });
      } else {
        // Build updated card with APPROVED status
        updatedCard = cards.buildFolderApprovalCardWithStatus({
          approvalId,
          requesterName,
          requesterEmail,
          folderName,
          justification: data.justification || '',
          status: 'approved',
          approverName: approver.name,
          permission: permission,
          duration: duration === 'permanent' ? 'Permanent' : duration,
          expiresAt: expiresAtFormatted,
          processedTime,
        });
      }
      
      // Send notification to requester (async, don't block)
      if (requesterId) {
        try {
          const channelService = getChannelService();
          if (channelService) {
            let notificationCard;
            
            if (isInvitationSent) {
              notificationCard = buildInvitationNotificationCard({
                recordTitle: folderName,
                itemType: 'folder',
                permission: permission,
                approverName: approver.name,
              });
            } else {
              notificationCard = cards.buildRequesterNotificationCard({
                approved: true,
                recordTitle: folderName,
                itemType: 'folder',
                permission: permission,
                duration: duration === 'permanent' ? 'Permanent' : duration,
                expiresAt: expiresAtFormatted,
                approverName: approver.name,
              });
            }
            
            channelService.sendDirectMessage(requesterId, {
              type: 'message',
              attachments: [{
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: notificationCard,
              }],
            }).then(sent => {
              if (sent) {
                log.debug(`Sent ${isInvitationSent ? 'invitation' : 'folder approval'} notification to requester`);
              }
            }).catch(err => {
              log.debug(`Could not send notification`, err.message);
            });
          }
        } catch (notifyError) {
          log.error('Error sending notification', notifyError.message);
        }
      }
      
      return { updatedCard };
    }
    
    case 'deny_record': {
      const { approvalId, recordTitle, requesterName, requesterId, requesterEmail } = data;
      const processedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
      
      log.debug('Denying record via Universal Action', { approver: approver.name, recordTitle, requesterName });
      
      // Store denial status for refresh mechanism
      storeApprovalStatus(approvalId, {
        status: 'denied',
        type: 'record',
        approverName: approver.name,
        requesterName,
        requesterEmail,
        recordTitle,
        justification: data.justification || '',
        processedTime,
      });
      
      // Build updated card with DENIED status
      const updatedCard = cards.buildRecordApprovalCardWithStatus({
        approvalId,
        requesterName,
        recordTitle,
        justification: data.justification || '',
        status: 'denied',
        approverName: approver.name,
        processedTime,
      });
      
      // Send notification to requester (but don't block on it)
      if (requesterId) {
        try {
          const channelService = getChannelService();
          if (channelService) {
            const notificationCard = cards.buildRequesterNotificationCard({
              approved: false,
              recordTitle: recordTitle,
              approverName: approver.name,
              denialReason: data.denialReason || null,
              itemType: 'record',
            });
            
            channelService.sendDirectMessage(requesterId, {
              type: 'message',
              attachments: [{
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: notificationCard,
              }],
            }).then(sent => {
              if (sent) {
                log.debug(`Sent denial notification to requester`);
              }
            }).catch(err => {
              log.debug(`Could not send notification`, err.message);
            });
          }
        } catch (notifyError) {
          log.error('Error sending notification', notifyError.message);
        }
      }
      
      return { updatedCard };
    }
    
    case 'deny_folder': {
      const { approvalId, folderName, requesterName, requesterId, requesterEmail } = data;
      const processedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
      
      log.debug('Denying folder via Universal Action', { approver: approver.name, folderName, requesterName });
      
      // Store denial status for refresh mechanism
      storeApprovalStatus(approvalId, {
        status: 'denied',
        type: 'folder',
        approverName: approver.name,
        requesterName,
        requesterEmail,
        folderName,
        justification: data.justification || '',
        processedTime,
      });
      
      // Build updated card with DENIED status
      const updatedCard = cards.buildFolderApprovalCardWithStatus({
        approvalId,
        requesterName,
        folderName,
        justification: data.justification || '',
        status: 'denied',
        approverName: approver.name,
        processedTime,
      });
      
      // Send notification to requester
      if (requesterId) {
        try {
          const channelService = getChannelService();
          if (channelService) {
            const notificationCard = cards.buildRequesterNotificationCard({
              approved: false,
              recordTitle: folderName,
              itemType: 'folder',
              approverName: approver.name,
              denialReason: data.denialReason || null,
            });
            
            channelService.sendDirectMessage(requesterId, {
              type: 'message',
              attachments: [{
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: notificationCard,
              }],
            }).then(sent => {
              if (sent) {
                log.debug(`Sent folder denial notification to requester`);
              }
            }).catch(err => {
              log.debug(`Could not send notification`, err.message);
            });
          }
        } catch (notifyError) {
          log.error('Error sending notification', notifyError.message);
        }
      }
      
      return { updatedCard };
    }
    
    case 'approve_share': {
      const { approvalId, recordUid, recordTitle, requesterName, requesterId, requesterEmail } = data;
      const duration = data.duration || '24h';
      const durationSeconds = parseDuration(duration) || 86400;
      const editable = data.editable === 'true' || data.editable === true;
      const processedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
      
      log.debug('Approving one-time share via Universal Action', { approver: approver.name, recordTitle, requesterName, duration, editable });
      
      if (!recordUid) {
        log.error('Missing record UID for share');
        return { error: 'Missing record UID' };
      }
      
      // Create the one-time share
      const result = await keeperClient.createOneTimeShare(
        recordUid,
        durationSeconds,
        editable
      );
      
      if (!result.success) {
        log.error('Failed to create one-time share', result.error);
        return { error: result.error };
      }
      
      // Format expiry date
      let expiresAtFormatted = 'N/A';
      if (result.expiresAt) {
        const expiryDate = new Date(result.expiresAt);
        expiresAtFormatted = expiryDate.toLocaleString('en-US', {
          month: '2-digit',
          day: '2-digit',
          year: 'numeric',
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit',
          hour12: false,
        });
      }
      
      // Store approval status for refresh mechanism
      storeApprovalStatus(approvalId, {
        status: 'approved',
        type: 'share',
        approverName: approver.name,
        requesterName,
        requesterEmail,
        recordTitle,
        justification: data.justification || '',
        duration,
        editable,
        expiresAt: expiresAtFormatted,
        shareUrl: result.shareUrl,
        processedTime,
      });
      
      // Build updated card with APPROVED status
      const updatedCard = cards.buildOneTimeShareApprovalCardWithStatus({
        approvalId,
        requesterName,
        requesterEmail,
        recordTitle,
        recordUid,
        justification: data.justification || '',
        status: 'approved',
        approverName: approver.name,
        duration,
        expiresAt: expiresAtFormatted,
        shareUrl: result.shareUrl,
        editable,
      });
      
      // Send share link to requester via DM
      if (requesterId) {
        try {
          const channelService = getChannelService();
          if (channelService) {
            // Build a card with the share link
            const shareLinkCard = {
              type: 'AdaptiveCard',
              '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
              version: '1.5',
              body: [
                {
                  type: 'TextBlock',
                  text: 'One-Time Share Link Created',
                  weight: 'Bolder',
                  size: 'Large',
                  color: 'Good',
                },
                {
                  type: 'TextBlock',
                  text: `Your one-time share request for **${recordTitle}** has been approved!`,
                  wrap: true,
                },
                {
                  type: 'FactSet',
                  facts: [
                    { title: 'Record', value: recordTitle },
                    { title: 'Duration', value: duration },
                    { title: 'Editable', value: editable ? 'Yes' : 'No' },
                    { title: 'Expires', value: expiresAtFormatted },
                    { title: 'Approved by', value: approver.name },
                  ],
                },
                {
                  type: 'Container',
                  separator: true,
                  spacing: 'Medium',
                  items: [
                    {
                      type: 'TextBlock',
                      text: 'Share Link:',
                      weight: 'Bolder',
                    },
                    {
                      type: 'TextBlock',
                      text: result.shareUrl,
                      wrap: true,
                      color: 'Accent',
                    },
                  ],
                },
                {
                  type: 'TextBlock',
                  text: '💡 This link can only be used once. Share it carefully.',
                  wrap: true,
                  isSubtle: true,
                  spacing: 'Medium',
                },
              ],
              actions: [
                {
                  type: 'Action.OpenUrl',
                  title: 'Open Share Link',
                  url: result.shareUrl,
                },
              ],
            };
            
            channelService.sendDirectMessage(requesterId, {
              type: 'message',
              attachments: [{
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: shareLinkCard,
              }],
            }).then(sent => {
              if (sent) {
                log.debug(`Sent share link to requester via DM`);
              }
            }).catch(err => {
              log.debug(`Could not send share link`, err.message);
            });
          }
        } catch (notifyError) {
          log.error('Error sending share link', notifyError.message);
        }
      }
      
      return { updatedCard };
    }
    
    case 'deny_share': {
      const { approvalId, recordUid, recordTitle, requesterName, requesterId, requesterEmail } = data;
      const processedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
      
      log.debug('Denying one-time share via Universal Action', { approver: approver.name, recordTitle, requesterName });
      
      // Store denial status for refresh mechanism
      storeApprovalStatus(approvalId, {
        status: 'denied',
        type: 'share',
        approverName: approver.name,
        requesterName,
        requesterEmail,
        recordTitle,
        justification: data.justification || '',
        processedTime,
      });
      
      // Build updated card with DENIED status
      const updatedCard = cards.buildOneTimeShareApprovalCardWithStatus({
        approvalId,
        requesterName,
        requesterEmail,
        recordTitle,
        recordUid,
        justification: data.justification || '',
        status: 'denied',
        approverName: approver.name,
      });
      
      // Send notification to requester
      if (requesterId) {
        try {
          const channelService = getChannelService();
          if (channelService) {
            const notificationCard = cards.buildRequesterNotificationCard({
              approved: false,
              recordTitle: recordTitle,
              itemType: 'one-time share',
              approverName: approver.name,
              denialReason: data.denialReason || null,
            });
            
            channelService.sendDirectMessage(requesterId, {
              type: 'message',
              attachments: [{
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: notificationCard,
              }],
            }).then(sent => {
              if (sent) {
                log.debug(`Sent share denial notification to requester`);
              }
            }).catch(err => {
              log.debug(`Could not send notification`, err.message);
            });
          }
        } catch (notifyError) {
          log.error('Error sending notification', notifyError.message);
        }
      }
      
      return { updatedCard };
    }
    
    default:
      // For other actions, fall back to the original handler
      await routeApprovalAction(context, data);
      return null;
  }
}

/**
 * Handle refresh action for approval cards
 * This is called when Teams auto-refreshes a card (via the refresh property)
 * Returns the appropriate card based on the current approval status
 * 
 * @param {Object} data - Refresh action data containing approvalId, type, and original card data
 * @returns {Object|null} - Updated card if status has changed, null otherwise
 */
async function handleRefreshApprovalCard(data) {
  const { approvalId, type } = data;
  
  if (!approvalId) {
    log.debug('Refresh: No approvalId provided');
    return null;
  }
  
  log.debug(`Checking refresh for approval ${approvalId} (type: ${type})`);
  
  const status = getApprovalStatus(approvalId);
  
  if (!status) {
    log.debug(`Approval ${approvalId} not yet processed, keeping original card`);
    return null;
  }
  
  log.debug(`Approval ${approvalId} has status: ${status.status}`);
  
  // Return the appropriate status card based on type
  if (type === 'record') {
    return cards.buildRecordApprovalCardWithStatus({
      approvalId: approvalId,
      requesterName: status.requesterName || data.requesterName,
      requesterEmail: status.requesterEmail || data.requesterEmail,
      recordTitle: status.recordTitle || data.recordTitle,
      justification: status.justification || data.justification,
      status: status.status, // 'approved' or 'denied'
      approverName: status.approverName,
      permission: status.permission,
      duration: status.duration,
      expiresAt: status.expiresAt,
      processedTime: status.processedTime,
    });
  } else if (type === 'folder') {
    return cards.buildFolderApprovalCardWithStatus({
      approvalId: approvalId,
      requesterName: status.requesterName || data.requesterName,
      requesterEmail: status.requesterEmail || data.requesterEmail,
      folderName: status.folderName || data.folderName,
      justification: status.justification || data.justification,
      status: status.status, // 'approved' or 'denied'
      approverName: status.approverName,
      permission: status.permission,
      duration: status.duration,
      expiresAt: status.expiresAt,
      processedTime: status.processedTime,
    });
  }
  
  log.debug(`Unknown type ${type} for refresh`);
  return null;
}

/**
 * Handle inline lookup actions (search from the card itself)
 * Called when user clicks "Search" button on the approval card
 * Now supports multiple results with dropdown selection
 * 
 * @param {string} verb - 'lookup_record', 'lookup_folder', or 'lookup_share'
 * @param {Object} data - Card action data
 * @param {string} searchQuery - The search query from the input field
 * @returns {Object} - Updated card with search results
 */
async function handleInlineLookup(verb, data, searchQuery) {
  const isFolder = verb === 'lookup_folder';
  const isShare = verb === 'lookup_share';
  const {
    approvalId,
    identifier,
    recordTitle,
    folderName,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    requesterName,
    justification,
  } = data;
  
  const query = searchQuery || identifier || (isFolder ? folderName : recordTitle) || '';
  
  const lookupType = isFolder ? 'folder' : (isShare ? 'share' : 'record');
  log.debug(`Inline ${lookupType} lookup`, { query, approvalId });
  
  if (!query.trim()) {
    // Return card with "no results" and allow retry
    if (isFolder) {
      return cards.buildFolderSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        noResults: true,
        originalFolderName: folderName,
      });
    } else if (isShare) {
      return cards.buildShareSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        noResults: true,
        originalRecordTitle: recordTitle,
      });
    } else {
      return cards.buildRecordSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        noResults: true,
        originalRecordTitle: recordTitle,
      });
    }
  }
  
  try {
    // Search using keeperClient - fetch up to 10 results
    // For shares, we search records (one-time-share is only for records)
    const results = isFolder
      ? await keeperClient.searchFolders(query, 10)
      : await keeperClient.searchRecords(query, 10);
    
    log.debug(`Search results for "${query}": ${results?.length || 0}`);
    
    if (!results || results.length === 0) {
      // No results found
      if (isFolder) {
        return cards.buildFolderSearchResultsCard({
          approvalId,
          requesterName,
          requesterId,
          requesterEmail,
          requesterAadObjectId,
          justification,
          identifier,
          searchQuery: query,
          noResults: true,
          originalFolderName: folderName,
        });
      } else if (isShare) {
        return cards.buildShareSearchResultsCard({
          approvalId,
          requesterName,
          requesterId,
          requesterEmail,
          requesterAadObjectId,
          justification,
          identifier,
          searchQuery: query,
          noResults: true,
          originalRecordTitle: recordTitle,
        });
      } else {
        return cards.buildRecordSearchResultsCard({
          approvalId,
          requesterName,
          requesterId,
          requesterEmail,
          requesterAadObjectId,
          justification,
          identifier,
          searchQuery: query,
          noResults: true,
          originalRecordTitle: recordTitle,
        });
      }
    }
    
    // Found results - pass all of them to the card builder
    if (isFolder) {
      const foundFolders = results.map(f => ({
        uid: f.uid || f.folder_uid,
        name: f.name || f.title || f.uid,
      }));
      
      return cards.buildFolderSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        foundFolders,
        originalFolderName: folderName,
      });
    } else if (isShare) {
      const foundRecords = results.map(r => ({
        uid: r.uid || r.record_uid,
        title: r.title || r.name || r.uid,
      }));
      
      return cards.buildShareSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        foundRecords,
        originalRecordTitle: recordTitle,
      });
    } else {
      const foundRecords = results.map(r => ({
        uid: r.uid || r.record_uid,
        title: r.title || r.name || r.uid,
      }));
      
      return cards.buildRecordSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        foundRecords,
        originalRecordTitle: recordTitle,
      });
    }
  } catch (error) {
    log.error(`Error searching ${lookupType}s`, error);
    
    // Return error card
    if (isFolder) {
      return cards.buildFolderSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        noResults: true,
        originalFolderName: folderName,
      });
    } else if (isShare) {
      return cards.buildShareSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        noResults: true,
        originalRecordTitle: recordTitle,
      });
    } else {
      return cards.buildRecordSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        noResults: true,
        originalRecordTitle: recordTitle,
      });
    }
  }
}

/**
 * Handle reset card actions - returns the original approval card
 * Called when user clicks "Reset" button to go back to the initial search state
 * 
 * @param {string} verb - 'reset_record_card', 'reset_folder_card', or 'reset_share_card'
 * @param {Object} data - Card action data
 * @returns {Object} - Original approval card
 */
function handleResetCard(verb, data) {
  const isFolder = verb === 'reset_folder_card';
  const isShare = verb === 'reset_share_card';
  const {
    approvalId,
    identifier,
    recordTitle,
    folderName,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    requesterName,
    justification,
  } = data;
  
  const resetType = isFolder ? 'folder' : (isShare ? 'share' : 'record');
  log.debug(`Resetting ${resetType} card`, { approvalId });
  
  if (isFolder) {
    return cards.buildFolderApprovalCard({
      approvalId,
      requesterName,
      requesterId,
      requesterEmail,
      requesterAadObjectId,
      folderName: folderName,
      folderUid: null, // No UID since this is a description-based request
      justification,
      isUid: false,
      identifier: identifier || folderName,
    });
  } else if (isShare) {
    return cards.buildOneTimeShareApprovalCard({
      approvalId,
      requesterName,
      requesterId,
      requesterEmail,
      requesterAadObjectId,
      recordTitle: recordTitle,
      recordUid: null, // No UID since this is a description-based request
      justification,
      isUid: false,
      identifier: identifier || recordTitle,
    });
  } else {
    return cards.buildRecordApprovalCard({
      approvalId,
      requesterName,
      requesterId,
      requesterEmail,
      requesterAadObjectId,
      recordTitle: recordTitle,
      recordUid: null, // No UID since this is a description-based request
      justification,
      isUid: false,
      identifier: identifier || recordTitle,
    });
  }
}

/**
 * Handle show_create_form action - returns inline create record form card
 * 
 * @param {Object} data - Card action data
 * @returns {Object} - Create record form card
 */
function handleShowCreateForm(data) {
  const {
    approvalId,
    identifier,
    recordTitle,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    requesterName,
    justification,
    searchQuery,
  } = data;
  
  log.debug('Showing create record form', { approvalId });
  
  return cards.buildRecordCreationCard({
    approvalId,
    requesterName,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    justification,
    identifier,
    originalRecordTitle: recordTitle,
    searchQuery: searchQuery || recordTitle,
  });
}

/**
 * Handle submit_create_record action - creates record and returns card with approval options
 * 
 * @param {Object} data - Card action data (approval context)
 * @param {Object} formData - Form input data
 * @returns {Object} - Created record card with approval options, or error card
 */
async function handleSubmitCreateRecord(data, formData) {
  const {
    approvalId,
    identifier,
    originalRecordTitle,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    requesterName,
    justification,
  } = data;
  
  const { recordTitle, recordLogin, recordPassword, recordUrl, recordNotes } = formData;
  
  log.debug('Creating record', { title: recordTitle, approvalId });
  
  // Validate required fields
  if (!recordTitle || !recordTitle.trim()) {
    return cards.buildRecordCreationCard({
      approvalId,
      requesterName,
      requesterId,
      requesterEmail,
      requesterAadObjectId,
      justification,
      identifier,
      originalRecordTitle,
      searchQuery: '',
      error: 'Title is required',
    });
  }
  
  if (!recordLogin || !recordLogin.trim()) {
    return cards.buildRecordCreationCard({
      approvalId,
      requesterName,
      requesterId,
      requesterEmail,
      requesterAadObjectId,
      justification,
      identifier,
      originalRecordTitle,
      searchQuery: recordTitle,
      error: 'Login is required',
    });
  }
  
  // Determine if we should generate password
  const generatePassword = !recordPassword || recordPassword.trim() === '' || recordPassword === '$GEN';
  const passwordToUse = generatePassword ? '$GEN' : recordPassword;
  
  // Create the record
  const result = await keeperClient.createRecord({
    title: recordTitle.trim(),
    login: recordLogin.trim(),
    password: passwordToUse,
    url: recordUrl?.trim() || null,
    notes: recordNotes?.trim() || null,
    generatePassword: generatePassword,
  });
  
  if (!result.success) {
    log.error('Failed to create record', result.error);
    // Return error card (show form again with error message)
    return {
      type: 'AdaptiveCard',
      '$schema': 'https://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.4',
      body: [
        { type: 'TextBlock', text: 'Record Creation Failed', weight: 'Bolder', size: 'Large', color: 'Attention' },
        { type: 'TextBlock', text: `Error: ${result.error || 'Unknown error'}`, wrap: true },
        { type: 'TextBlock', text: 'Please try again or contact support.', wrap: true, isSubtle: true },
      ],
      actions: [
        {
          type: 'Action.Execute',
          title: 'Try Again',
          verb: 'show_create_form',
          data: { action: 'show_create_form', approvalId: approvalId || '', identifier: identifier || '', recordTitle: originalRecordTitle || '', requesterId: requesterId || '', requesterEmail: requesterEmail || '', requesterAadObjectId: requesterAadObjectId || '', requesterName: requesterName || '', justification: justification || '' },
        },
        {
          type: 'Action.Execute',
          title: 'Cancel',
          verb: 'cancel_create_form',
          data: { action: 'cancel_create_form', approvalId: approvalId || '', identifier: identifier || '', recordTitle: originalRecordTitle || '', requesterId: requesterId || '', requesterEmail: requesterEmail || '', requesterAadObjectId: requesterAadObjectId || '', requesterName: requesterName || '', justification: justification || '' },
        },
      ],
    };
  }
  
  log.debug('Record created successfully', result.recordUid);
  
  // Safe values for the card
  const safeApprovalId = approvalId || '';
  const safeRecordUid = result.recordUid;
  const safeRecordTitle = recordTitle.trim();
  const safeRequesterId = requesterId || '';
  const safeRequesterEmail = requesterEmail || '';
  const safeRequesterName = requesterName || 'Unknown';
  const safeJustification = justification || '';
  
  // Build a simple card matching the exact format of working cards
  // Use http:// schema and minimal actions (2 only)
  const createdRecordCard = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: [
      { type: 'TextBlock', text: 'Record Created Successfully!', weight: 'Bolder', size: 'Large' },
      { type: 'TextBlock', text: `Requester: ${safeRequesterName}`, wrap: true },
      { type: 'TextBlock', text: `Record: ${safeRecordTitle}`, wrap: true, weight: 'Bolder' },
      { type: 'TextBlock', text: `UID: ${safeRecordUid}`, wrap: true, size: 'Small' },
    ],
    actions: [
      {
        type: 'Action.Execute',
        title: 'Approve (View Only, 1h)',
        style: 'positive',
        verb: 'approve_record',
        data: { 
          action: 'approve_record', 
          approvalId: safeApprovalId, 
          recordUid: safeRecordUid, 
          recordTitle: safeRecordTitle, 
          requesterId: safeRequesterId, 
          requesterEmail: safeRequesterEmail, 
          requesterName: safeRequesterName, 
          justification: safeJustification,
          permission: 'view_only',
          duration: '1h',
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny',
        style: 'destructive',
        verb: 'deny_record',
        data: { 
          action: 'deny_record', 
          approvalId: safeApprovalId, 
          recordUid: safeRecordUid, 
          recordTitle: safeRecordTitle, 
          requesterId: safeRequesterId, 
          requesterEmail: safeRequesterEmail, 
          requesterName: safeRequesterName, 
          justification: safeJustification,
        },
      },
    ],
  };
  
  log.debug('Returning created record card');
  return createdRecordCard;
}

/**
 * Handle cancel_create_form action - returns to search results card with re-fetched results
 * 
 * @param {Object} data - Card action data
 * @returns {Object} - Search results card with actual search results
 */
async function handleCancelCreateForm(data) {
  const {
    approvalId,
    identifier,
    recordTitle,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    requesterName,
    justification,
    searchQuery,
  } = data;
  
  const query = searchQuery || recordTitle || identifier;
  log.debug('Cancelling create form, re-running search', { approvalId, query });
  
  // Re-run the search to restore previous results
  let searchResults = [];
  let noResults = true;
  
  if (query) {
    try {
      searchResults = await keeperClient.searchRecords(query);
      noResults = !searchResults || searchResults.length === 0;
      log.debug(`Search results for "${query}": ${searchResults.length}`);
    } catch (error) {
      log.error('Error re-running search', error);
    }
  }
  
  return cards.buildRecordSearchResultsCard({
    approvalId,
    requesterName,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    justification,
    identifier,
    searchQuery: query,
    records: searchResults,
    noResults,
    originalRecordTitle: recordTitle,
  });
}

module.exports = {
  routeApprovalAction,
  routeApprovalActionWithCardResponse,
  handleRecordApproval,
  handleRecordDenial,
  handleFolderApproval,
  handleFolderDenial,
  handleShareApproval,
  handleShareDenial,
  handleRefreshApprovalCard,
  handleInlineLookup,
  handleResetCard,
  handleShowCreateForm,
  handleSubmitCreateRecord,
  handleCancelCreateForm,
  parseDuration,
  DURATION_MAP,
};
