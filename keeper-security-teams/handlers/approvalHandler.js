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
const { getChannelService, getApprovalActivityId, removeApprovalActivityId } = require('../services');

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
    console.log(`[ApprovalHandler] Using replyToId as fallback: ${activity.replyToId}`);
    activityId = activity.replyToId;
  }
  
  if (!activityId) {
    console.log(`[ApprovalHandler] No activity ID found for approval ${approvalId}`);
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

  console.log(`[ApprovalHandler] Updating approval ${approvalId}:`, {
    activityId,
    conversationId: conversationId.substring(0, 50) + '...',
    serviceUrl: serviceUrl.substring(0, 50) + '...',
  });

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
      console.log('[ApprovalHandler] Updated via context.updateActivity');
      removeApprovalActivityId(approvalId);
      return true;
    } catch (e) {
      console.log('[ApprovalHandler] context.updateActivity failed:', e.message);
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
      
      console.log('[ApprovalHandler] Updated via ApiClient with context serviceUrl');
      removeApprovalActivityId(approvalId);
      return true;
    } catch (e) {
      console.log('[ApprovalHandler] ApiClient update failed:', e.message);
    }
  }

  // Method 3: Try using the channel service's updateApprovalCard as last resort
  if (channelService) {
    try {
      const success = await channelService.updateApprovalCard(activityId, updatedCard);
      if (success) {
        removeApprovalActivityId(approvalId);
        console.log('[ApprovalHandler] Updated via channelService');
        return true;
      }
    } catch (e) {
      console.log('[ApprovalHandler] channelService.updateApprovalCard failed:', e.message);
    }
  }

  throw new Error('All update methods failed');
}

/**
 * Handle approval of a record access request
 */
async function handleRecordApproval(context, data) {
  console.log('[ApprovalHandler] handleRecordApproval called with data:', JSON.stringify(data));
  
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
    await context.send('❌ Error: Missing record UID');
    return;
  }
  
  if (!requesterEmail) {
    await context.send('❌ Error: Missing requester email. Cannot grant access without email.');
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
    // Format expiry date
    let expiresAtFormatted = 'Permanent';
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
    
    // Build the updated card with APPROVED status and all details
    const updatedCard = cards.buildRecordApprovalCardWithStatus({
      approvalId: approvalId,
      requesterName: requesterName,
      requesterEmail: requesterEmail,
      recordTitle: recordTitle,
      justification: justification,
      status: 'approved',
      approverName: approver.name,
      permission: permission,
      duration: duration === 'permanent' ? 'Permanent' : duration,
      expiresAt: duration === 'permanent' ? null : expiresAtFormatted,
    });
    
    // Try to update the original approval card using stored activity ID
    try {
      await tryUpdateApprovalCard(approvalId, updatedCard, context);
      console.log('[ApprovalHandler] Successfully updated approval card with APPROVED status');
    } catch (updateError) {
      console.log('[ApprovalHandler] Failed to update activity, sending new message:', updateError.message);
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
          const notificationCard = cards.buildRequesterNotificationCard({
            approved: true,
            recordTitle: recordTitle,
            permission: permission,
            duration: duration === 'permanent' ? 'Permanent' : duration,
            expiresAt: duration === 'permanent' ? null : expiresAtFormatted,
            approverName: approver.name,
          });
          
          const notificationSent = await channelService.sendDirectMessage(requesterId, {
            type: 'message',
            attachments: [{
              contentType: 'application/vnd.microsoft.card.adaptive',
              content: notificationCard,
            }],
          });
          
          if (notificationSent) {
            console.log(`[ApprovalHandler] Sent approval notification to requester: ${requesterId}`);
          } else {
            console.log(`[ApprovalHandler] Could not send notification to requester (no reference stored)`);
          }
        }
      } catch (notifyError) {
        console.error('[ApprovalHandler] Error sending requester notification:', notifyError.message);
      }
    }
  } else {
    await context.send('❌ Failed to grant access: ' + result.error);
  }
}

/**
 * Handle denial of a record access request
 */
async function handleRecordDenial(context, data) {
  console.log('[ApprovalHandler] handleRecordDenial called with data:', JSON.stringify(data));
  
  const approver = getApproverInfo(context.activity);
  const recordTitle = data.recordTitle || data.recordUid || 'Unknown Record';
  const requesterName = data.requesterName || 'User';
  const approvalId = data.approvalId || 'N/A';
  const justification = data.justification || '';
  
  console.log('[ApprovalHandler] Denying record access:', { approver: approver.name, recordTitle, requesterName });
  
  // Build updated card with DENIED status
  const updatedCard = cards.buildRecordApprovalCardWithStatus({
    approvalId,
    requesterName,
    recordTitle,
    justification,
    status: 'denied',
    approverName: approver.name,
  });
  
  // Try to update the original approval card using stored activity ID
  try {
    await tryUpdateApprovalCard(approvalId, updatedCard, context);
    console.log('[ApprovalHandler] Updated original card with denied status');
  } catch (error) {
    console.error('[ApprovalHandler] Failed to update card, sending new message:', error.message);
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
          console.log(`[ApprovalHandler] Sent denial notification to requester: ${requesterId}`);
        } else {
          console.log(`[ApprovalHandler] Could not send notification to requester (no reference stored)`);
        }
      }
    } catch (notifyError) {
      console.error('[ApprovalHandler] Error sending requester notification:', notifyError.message);
    }
  }
  
  console.log('[ApprovalHandler] Denial complete');
}

/**
 * Handle approval of a folder access request
 */
async function handleFolderApproval(context, data) {
  console.log('[ApprovalHandler] handleFolderApproval called with data:', JSON.stringify(data));
  
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
    await context.send('❌ Error: Missing folder UID');
    return;
  }
  
  if (!requesterEmail) {
    await context.send('❌ Error: Missing requester email. Cannot grant access without email.');
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
    // Format expiry date
    let expiresAtFormatted = 'Permanent';
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
    
    // Build the updated card with APPROVED status
    const updatedCard = cards.buildFolderApprovalCardWithStatus({
      approvalId: approvalId,
      requesterName: requesterName,
      requesterEmail: requesterEmail,
      folderName: folderName,
      justification: justification,
      status: 'approved',
      approverName: approver.name,
      permission: permission,
      duration: duration === 'permanent' ? 'Permanent' : duration,
      expiresAt: duration === 'permanent' ? null : expiresAtFormatted,
    });
    
    // Try to update the original approval card using stored activity ID
    try {
      await tryUpdateApprovalCard(approvalId, updatedCard, context);
      console.log('[ApprovalHandler] Successfully updated folder approval card with APPROVED status');
    } catch (updateError) {
      console.log('[ApprovalHandler] Failed to update activity, sending new message:', updateError.message);
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
          const notificationCard = cards.buildRequesterNotificationCard({
            approved: true,
            recordTitle: folderName,
            itemType: 'folder',
            permission: permission,
            duration: duration === 'permanent' ? 'Permanent' : duration,
            expiresAt: duration === 'permanent' ? null : expiresAtFormatted,
            approverName: approver.name,
          });
          
          const notificationSent = await channelService.sendDirectMessage(requesterId, {
            type: 'message',
            attachments: [{
              contentType: 'application/vnd.microsoft.card.adaptive',
              content: notificationCard,
            }],
          });
          
          if (notificationSent) {
            console.log(`[ApprovalHandler] Sent folder approval notification to requester: ${requesterId}`);
          } else {
            console.log(`[ApprovalHandler] Could not send notification to requester (no reference stored)`);
          }
        }
      } catch (notifyError) {
        console.error('[ApprovalHandler] Error sending requester notification:', notifyError.message);
      }
    }
  } else {
    await context.send('❌ Failed to grant folder access: ' + result.error);
  }
}

/**
 * Handle denial of a folder access request
 */
async function handleFolderDenial(context, data) {
  console.log('[ApprovalHandler] handleFolderDenial called with data:', JSON.stringify(data));
  
  const approver = getApproverInfo(context.activity);
  const folderName = data.folderName || data.folderUid || 'Unknown Folder';
  const requesterName = data.requesterName || 'User';
  const approvalId = data.approvalId || 'N/A';
  const justification = data.justification || '';
  
  console.log('[ApprovalHandler] Denying folder access:', { approver: approver.name, folderName, requesterName });
  
  // Build updated card with DENIED status
  const updatedCard = cards.buildFolderApprovalCardWithStatus({
    approvalId,
    requesterName,
    folderName,
    justification,
    status: 'denied',
    approverName: approver.name,
  });
  
  // Try to update the original approval card using stored activity ID
  try {
    await tryUpdateApprovalCard(approvalId, updatedCard, context);
    console.log('[ApprovalHandler] Updated original folder card with denied status');
  } catch (error) {
    console.error('[ApprovalHandler] Failed to update folder card, sending new message:', error.message);
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
          console.log(`[ApprovalHandler] Sent folder denial notification to requester: ${requesterId}`);
        } else {
          console.log(`[ApprovalHandler] Could not send notification to requester (no reference stored)`);
        }
      }
    } catch (notifyError) {
      console.error('[ApprovalHandler] Error sending requester notification:', notifyError.message);
    }
  }
  
  console.log('[ApprovalHandler] Folder denial complete');
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
    await context.send('❌ Error: Missing record UID');
    return;
  }
  
  // Create the share
  await context.send('🔗 Creating one-time share link...');
  
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
      text: '✅ Share link created for ' + requesterName,
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: shareCard,
      }],
    });
  } else {
    await context.send('❌ Failed to create share link: ' + result.error);
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
  
  await context.send('❌ Share request denied for ' + requesterName + ' for record **' + recordTitle + '**');
}

/**
 * Handle search_records action - Opens task module
 * In Teams, task modules are opened via invoke responses
 * @param {Object} context - Teams turn context
 * @param {Object} data - Card action data
 */
async function handleSearchRecordsAction(context, data) {
  const { identifier, approvalId, requesterEmail, requesterName, requesterId, requesterAadObjectId, justification } = data;
  
  console.log('[ApprovalHandler] Search records action:', { identifier, approvalId });
  
  // Store approval context for later use (when record is selected)
  const approvalContext = {
    approvalId,
    requesterEmail,
    requesterName,
    requesterId,
    requesterAadObjectId,
    justification,
    identifier,
    searchType: 'record',
  };
  
  try {
    const { handleTaskFetch } = require('./taskModuleHandler');
    
    // Create a task module request
    const taskModuleRequest = {
      value: {
        data: {
          type: 'search-record',
          query: identifier,
          approvalId: approvalId,
          approvalContext: approvalContext,
        },
      },
    };
    
    // Get the task module response
    const taskModuleResponse = await handleTaskFetch(context, taskModuleRequest);
    
    // Return the response - this will be sent as invoke response
    // The SDK should handle converting this to proper invoke response format
    return taskModuleResponse;
    
  } catch (error) {
    console.error('[ApprovalHandler] Error opening task module:', error);
    // Can't send message here as we need to return invoke response
    // The error will be logged and task module won't open
    throw error;
  }
}

/**
 * Handle search_folders action - Opens task module for folder search
 * In Teams, task modules are opened via invoke responses
 * @param {Object} context - Teams turn context
 * @param {Object} data - Card action data
 */
async function handleSearchFoldersAction(context, data) {
  const { identifier, approvalId, requesterEmail, requesterName, requesterId, requesterAadObjectId, justification, folderName } = data;
  
  console.log('[ApprovalHandler] Search folders action:', { identifier, approvalId });
  
  // Store approval context for later use (when folder is selected)
  const approvalContext = {
    approvalId,
    requesterEmail,
    requesterName,
    requesterId,
    requesterAadObjectId,
    justification,
    identifier,
    folderName,
    searchType: 'folder',
  };
  
  try {
    const { handleTaskFetch } = require('./taskModuleHandler');
    
    // Create a task module request
    const taskModuleRequest = {
      value: {
        data: {
          type: 'search-folder',
          action: 'search_folders',
          query: identifier,
          approvalId: approvalId,
          approvalContext: approvalContext,
        },
      },
    };
    
    // Get the task module response
    const taskModuleResponse = await handleTaskFetch(context, taskModuleRequest);
    
    // Return the response - this will be sent as invoke response
    return taskModuleResponse;
    
  } catch (error) {
    console.error('[ApprovalHandler] Error opening folder search task module:', error);
    throw error;
  }
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
  
  console.log('[ApprovalHandler] routeApprovalActionWithCardResponse:', action);
  
  switch (action) {
    case 'approve_record': {
      const { approvalId, recordUid, recordTitle, requesterName, requesterId, requesterEmail } = data;
      const permission = data.permission || 'view_only';
      const duration = data.duration || '24h';
      const durationSeconds = parseDuration(duration);
      
      console.log('[ApprovalHandler] Approving record via Universal Action:', { 
        approver: approver.name, 
        recordTitle, 
        requesterName,
        permission,
        duration
      });
      
      if (!recordUid) {
        console.error('[ApprovalHandler] Missing record UID');
        return { error: 'Missing record UID' };
      }
      
      if (!requesterEmail) {
        console.error('[ApprovalHandler] Missing requester email');
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
        console.error('[ApprovalHandler] Failed to grant access:', result.error);
        return { error: result.error };
      }
      
      // Format expiry date
      let expiresAtFormatted = 'Permanent';
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
      
      // Build updated card with APPROVED status
      const updatedCard = cards.buildRecordApprovalCardWithStatus({
        approvalId,
        requesterName,
        requesterEmail,
        recordTitle,
        justification: data.justification || '',
        status: 'approved',
        approverName: approver.name,
        permission: permission,
        duration: duration === 'permanent' ? 'Permanent' : duration,
        expiresAt: duration === 'permanent' ? null : expiresAtFormatted,
      });
      
      // Send notification to requester (async, don't block)
      if (requesterId) {
        try {
          const channelService = getChannelService();
          if (channelService) {
            const notificationCard = cards.buildRequesterNotificationCard({
              approved: true,
              recordTitle: recordTitle,
              permission: permission,
              duration: duration === 'permanent' ? 'Permanent' : duration,
              expiresAt: duration === 'permanent' ? null : expiresAtFormatted,
              approverName: approver.name,
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
                console.log(`[ApprovalHandler] Sent approval notification to requester`);
              }
            }).catch(err => {
              console.log(`[ApprovalHandler] Could not send notification:`, err.message);
            });
          }
        } catch (notifyError) {
          console.error('[ApprovalHandler] Error sending notification:', notifyError.message);
        }
      }
      
      return { updatedCard };
    }
    
    case 'approve_folder': {
      const { approvalId, folderUid, folderName, requesterName, requesterId, requesterEmail } = data;
      const permission = data.permission || 'no_permissions';
      const duration = data.duration || '24h';
      const durationSeconds = parseDuration(duration);
      
      console.log('[ApprovalHandler] Approving folder via Universal Action:', { 
        approver: approver.name, 
        folderName, 
        requesterName,
        permission,
        duration
      });
      
      if (!folderUid) {
        console.error('[ApprovalHandler] Missing folder UID');
        return { error: 'Missing folder UID' };
      }
      
      if (!requesterEmail) {
        console.error('[ApprovalHandler] Missing requester email');
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
        console.error('[ApprovalHandler] Failed to grant folder access:', result.error);
        return { error: result.error };
      }
      
      // Format expiry date
      let expiresAtFormatted = 'Permanent';
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
      
      // Build updated card with APPROVED status
      const updatedCard = cards.buildFolderApprovalCardWithStatus({
        approvalId,
        requesterName,
        requesterEmail,
        folderName,
        justification: data.justification || '',
        status: 'approved',
        approverName: approver.name,
        permission: permission,
        duration: duration === 'permanent' ? 'Permanent' : duration,
        expiresAt: duration === 'permanent' ? null : expiresAtFormatted,
      });
      
      // Send notification to requester (async, don't block)
      if (requesterId) {
        try {
          const channelService = getChannelService();
          if (channelService) {
            const notificationCard = cards.buildRequesterNotificationCard({
              approved: true,
              recordTitle: folderName,
              itemType: 'folder',
              permission: permission,
              duration: duration === 'permanent' ? 'Permanent' : duration,
              expiresAt: duration === 'permanent' ? null : expiresAtFormatted,
              approverName: approver.name,
            });
            
            channelService.sendDirectMessage(requesterId, {
              type: 'message',
              attachments: [{
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: notificationCard,
              }],
            }).then(sent => {
              if (sent) {
                console.log(`[ApprovalHandler] Sent folder approval notification to requester`);
              }
            }).catch(err => {
              console.log(`[ApprovalHandler] Could not send notification:`, err.message);
            });
          }
        } catch (notifyError) {
          console.error('[ApprovalHandler] Error sending notification:', notifyError.message);
        }
      }
      
      return { updatedCard };
    }
    
    case 'deny_record': {
      const { approvalId, recordTitle, requesterName, requesterId } = data;
      
      console.log('[ApprovalHandler] Denying record via Universal Action:', { 
        approver: approver.name, 
        recordTitle, 
        requesterName 
      });
      
      // Build updated card with DENIED status
      const updatedCard = cards.buildRecordApprovalCardWithStatus({
        approvalId,
        requesterName,
        recordTitle,
        justification: data.justification || '',
        status: 'denied',
        approverName: approver.name,
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
                console.log(`[ApprovalHandler] Sent denial notification to requester`);
              }
            }).catch(err => {
              console.log(`[ApprovalHandler] Could not send notification:`, err.message);
            });
          }
        } catch (notifyError) {
          console.error('[ApprovalHandler] Error sending notification:', notifyError.message);
        }
      }
      
      // Return the updated card - Teams will replace the original card with this
      return { updatedCard };
    }
    
    case 'deny_folder': {
      const { approvalId, folderName, requesterName, requesterId } = data;
      
      console.log('[ApprovalHandler] Denying folder via Universal Action:', { 
        approver: approver.name, 
        folderName, 
        requesterName 
      });
      
      // Build updated card with DENIED status
      const updatedCard = cards.buildFolderApprovalCardWithStatus({
        approvalId,
        requesterName,
        folderName,
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
                console.log(`[ApprovalHandler] Sent folder denial notification to requester`);
              }
            }).catch(err => {
              console.log(`[ApprovalHandler] Could not send notification:`, err.message);
            });
          }
        } catch (notifyError) {
          console.error('[ApprovalHandler] Error sending notification:', notifyError.message);
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

module.exports = {
  routeApprovalAction,
  routeApprovalActionWithCardResponse,
  handleRecordApproval,
  handleRecordDenial,
  handleFolderApproval,
  handleFolderDenial,
  handleShareApproval,
  handleShareDenial,
  handleSearchRecordsAction,
  handleSearchFoldersAction,
  parseDuration,
  DURATION_MAP,
};
