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
 * Handle approval of a record access request
 */
async function handleRecordApproval(context, data) {
  const approver = getApproverInfo(context.activity);
  const permission = data.permission || 'view_only';
  const duration = data.duration || '24h';
  const durationSeconds = parseDuration(duration);
  
  const recordUid = data.recordUid;
  const recordTitle = data.recordTitle || recordUid;
  const requesterEmail = data.requesterEmail;
  const requesterName = data.requesterName || 'User';
  
  if (!recordUid) {
    await context.send('❌ Error: Missing record UID');
    return;
  }
  
  if (!requesterEmail) {
    await context.send('❌ Error: Missing requester email. Cannot grant access without email.');
    return;
  }
  
  // Grant access
  await context.send('🔄 Granting access...');
  
  const result = await keeperClient.grantRecordAccess(
    recordUid,
    requesterEmail,
    permission,
    durationSeconds
  );
  
  if (result.success) {
    // Update the card to show approved
    const approvedCard = cards.buildApprovedMessageCard(
      approver.name,
      permission,
      duration === 'permanent' ? 'Permanent' : duration,
      recordTitle,
      'record'
    );
    
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: approvedCard,
      }],
    });
    
    // Send result notification
    const resultCard = cards.buildApprovalResultCard({
      approved: true,
      approverName: approver.name,
      itemName: recordTitle,
      itemType: 'record',
      permission: permission,
      duration: duration === 'permanent' ? 'Permanent' : duration,
      expiresAt: result.expiresAt,
    });
    
    await context.send({
      type: 'message',
      text: '✅ Access granted to ' + requesterName + ' for record **' + recordTitle + '**',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: resultCard,
      }],
    });
  } else {
    await context.send('❌ Failed to grant access: ' + result.error);
  }
}

/**
 * Handle denial of a record access request
 */
async function handleRecordDenial(context, data) {
  const approver = getApproverInfo(context.activity);
  const recordTitle = data.recordTitle || data.recordUid;
  const requesterName = data.requesterName || 'User';
  
  // Update the card to show denied
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
  
  await context.send('❌ Access request denied for ' + requesterName + ' to record **' + recordTitle + '**');
}

/**
 * Handle approval of a folder access request
 */
async function handleFolderApproval(context, data) {
  const approver = getApproverInfo(context.activity);
  const permission = data.permission || 'no_permissions';
  const duration = data.duration || '24h';
  const durationSeconds = parseDuration(duration);
  
  const folderUid = data.folderUid;
  const folderName = data.folderName || folderUid;
  const requesterEmail = data.requesterEmail;
  const requesterName = data.requesterName || 'User';
  
  if (!folderUid) {
    await context.send('❌ Error: Missing folder UID');
    return;
  }
  
  if (!requesterEmail) {
    await context.send('❌ Error: Missing requester email. Cannot grant access without email.');
    return;
  }
  
  // Grant access
  await context.send('🔄 Granting folder access...');
  
  const result = await keeperClient.grantFolderAccess(
    folderUid,
    requesterEmail,
    permission,
    durationSeconds
  );
  
  if (result.success) {
    const approvedCard = cards.buildApprovedMessageCard(
      approver.name,
      permission,
      duration === 'permanent' ? 'Permanent' : duration,
      folderName,
      'folder'
    );
    
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: approvedCard,
      }],
    });
    
    const resultCard = cards.buildApprovalResultCard({
      approved: true,
      approverName: approver.name,
      itemName: folderName,
      itemType: 'folder',
      permission: permission,
      duration: duration === 'permanent' ? 'Permanent' : duration,
      expiresAt: result.expiresAt,
    });
    
    await context.send({
      type: 'message',
      text: '✅ Access granted to ' + requesterName + ' for folder **' + folderName + '**',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: resultCard,
      }],
    });
  } else {
    await context.send('❌ Failed to grant folder access: ' + result.error);
  }
}

/**
 * Handle denial of a folder access request
 */
async function handleFolderDenial(context, data) {
  const approver = getApproverInfo(context.activity);
  const folderName = data.folderName || data.folderUid;
  const requesterName = data.requesterName || 'User';
  
  const deniedCard = cards.buildDeniedMessageCard(
    approver.name,
    null,
    folderName,
    'folder'
  );
  
  await context.send({
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: deniedCard,
    }],
  });
  
  await context.send('❌ Access request denied for ' + requesterName + ' to folder **' + folderName + '**');
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
  const { identifier, approvalId, requesterEmail, requesterName, requesterId, justification } = data;
  
  console.log('[ApprovalHandler] Search records action:', { identifier, approvalId });
  
  // Store approval context for later use (when record is selected)
  const approvalContext = {
    approvalId,
    requesterEmail,
    requesterName,
    requesterId,
    justification,
    identifier,
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

module.exports = {
  routeApprovalAction,
  handleRecordApproval,
  handleRecordDenial,
  handleFolderApproval,
  handleFolderDenial,
  handleShareApproval,
  handleShareDenial,
  handleSearchRecordsAction,
  parseDuration,
  DURATION_MAP,
};
