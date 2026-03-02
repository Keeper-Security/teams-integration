/**
 * Approval Router
 * 
 * Routes card actions to appropriate handlers.
 */

const keeperClient = require('../../services/keeperClient');
const cards = require('../../cards');
const { getChannelService, getApprovalStatus, storeApprovalStatus, createLogger } = require('../../services');
const { isPermissionConflictError, isRecordOwnerError } = require('../../utils/helpers');
const { 
  DURATION_MAP,
  parseDuration, 
  getDisplayDuration,
  formatExpiryDate, 
  getApproverInfo, 
  tryUpdateApprovalCard,
  buildInvitationNotificationCard,
} = require('./helpers');
const { handleRecordApproval, handleRecordDenial } = require('./recordHandler');
const { handleFolderApproval, handleFolderDenial } = require('./folderHandler');
const { handleShareApproval, handleShareDenial } = require('./shareHandler');

const log = createLogger('ApprovalRouter');

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
  
  if (approvalId) {
    const existingStatus = getApprovalStatus(approvalId);
    if (existingStatus) {
      log.debug(`Approval ${approvalId} already processed: ${existingStatus.status}`);
      
      const statusText = existingStatus.status === 'approved' ? 'APPROVED' : 'DENIED';
      const itemName = existingStatus.recordTitle || existingStatus.folderName || 'the requested item';
      const itemType = existingStatus.type === 'folder' ? 'Folder' : 'Record';
      
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
    case 'approve_record':
      return handleRecordApprovalWithCardResponse(context, data, approver);
      
    case 'approve_folder':
      return handleFolderApprovalWithCardResponse(context, data, approver);
      
    case 'deny_record':
      return handleRecordDenialWithCardResponse(context, data, approver);
      
    case 'deny_folder':
      return handleFolderDenialWithCardResponse(context, data, approver);
      
    case 'approve_share':
      return handleShareApprovalWithCardResponse(context, data, approver);
      
    case 'deny_share':
      return handleShareDenialWithCardResponse(context, data, approver);
      
    default:
      await routeApprovalAction(context, data);
      return null;
  }
}

async function handleRecordApprovalWithCardResponse(context, data, approver) {
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
  
  const result = await keeperClient.grantRecordAccess(
    recordUid,
    requesterEmail,
    permission,
    durationSeconds
  );
  
  if (!result.success) {
    if (result.alreadyHasAccess) {
      log.info('User already has access to record', { recordUid, requesterEmail, currentPermission: result.currentPermission });
      
      const alreadyHasAccessCard = cards.buildRecordAlreadyHasAccessCard({
        approvalId,
        requesterName,
        requesterEmail,
        recordTitle,
        recordUid,
        justification: data.justification || '',
        currentPermission: result.currentPermission,
        currentPermissionLabel: result.currentPermissionLabel,
        approverName: approver.name,
        processedTime,
      });
      
      // Non-blocking update to avoid Teams timeout
      tryUpdateApprovalCard(approvalId, alreadyHasAccessCard, context)
        .then(() => log.debug('Updated channel card with already has access status'))
        .catch(err => log.debug('Could not update channel card', err.message));
      
      return { updatedCard: alreadyHasAccessCard };
    }
    
    if (result.isOwnerError || isRecordOwnerError(result.error)) {
      log.info('Record owner error detected in Universal Action', { recordUid, requesterEmail });
      
      const ownerStatusCard = cards.buildRecordApprovalCardWithStatus({
        approvalId,
        requesterName,
        requesterEmail,
        recordTitle,
        justification: data.justification || '',
        status: 'owner',
        statusMessage: 'User Already Has Full Access (Owner)',
        approverName: approver.name,
        permission,
        processedTime,
      });
      
      // Non-blocking update to avoid Teams timeout
      tryUpdateApprovalCard(approvalId, ownerStatusCard, context)
        .then(() => log.debug('Updated channel card with owner status'))
        .catch(err => log.debug('Could not update channel card', err.message));
      
      return { updatedCard: ownerStatusCard };
    }
    
    if (isPermissionConflictError(result.error)) {
      log.info('Permission conflict detected in Universal Action', { recordUid, requesterEmail, error: result.error });
      return { error: result.error, keepCardActive: true };
    }
    
    log.error('Failed to grant access', result.error);
    return { error: result.error };
  }
  
  const expiresAtFormatted = formatExpiryDate(result.expiresAt);
  const isInvitationSent = result.invitationSent;
  
  storeApprovalStatus(approvalId, {
    status: isInvitationSent ? 'invitation_sent' : 'approved',
    type: 'record',
    approverName: approver.name,
    requesterName,
    requesterEmail,
    recordTitle,
    justification: data.justification || '',
    permission,
    duration: getDisplayDuration(permission, duration, 'record'),
    expiresAt: expiresAtFormatted,
    processedTime,
    invitationSent: isInvitationSent,
  });
  
  let updatedCard;
  
  if (isInvitationSent) {
    log.info('Share invitation sent for record (user has no Keeper account)');
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
    updatedCard = cards.buildRecordApprovalCardWithStatus({
      approvalId,
      requesterName,
      requesterEmail,
      recordTitle,
      justification: data.justification || '',
      status: 'approved',
      approverName: approver.name,
      permission: permission,
      duration: getDisplayDuration(permission, duration, 'record'),
      expiresAt: expiresAtFormatted,
      processedTime,
    });
  }
  
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
            duration: getDisplayDuration(permission, duration, 'record'),
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
  
  // Non-blocking update to avoid Teams timeout
  tryUpdateApprovalCard(approvalId, updatedCard, context)
    .then(() => log.debug('Updated channel card with approved status'))
    .catch(err => log.debug('Could not update channel card', err.message));
  
  return { updatedCard };
}

async function handleFolderApprovalWithCardResponse(context, data, approver) {
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
  
  const result = await keeperClient.grantFolderAccess(
    folderUid,
    requesterEmail,
    permission,
    durationSeconds
  );
  
  if (!result.success) {
    if (result.alreadyHasAccess) {
      log.info('User already has access to folder', { folderUid, requesterEmail, currentPermission: result.currentPermission });
      
      const alreadyHasAccessCard = cards.buildFolderAlreadyHasAccessCard({
        approvalId,
        requesterName,
        requesterEmail,
        folderName,
        folderUid,
        justification: data.justification || '',
        currentPermission: result.currentPermission,
        currentPermissionLabel: result.currentPermissionLabel,
        approverName: approver.name,
        processedTime,
      });
      
      // Non-blocking update to avoid Teams timeout
      tryUpdateApprovalCard(approvalId, alreadyHasAccessCard, context)
        .then(() => log.debug('Updated channel card with already has access status'))
        .catch(err => log.debug('Could not update channel card', err.message));
      
      return { updatedCard: alreadyHasAccessCard };
    }
    
    if (result.isFullAccessError) {
      log.info('Folder full access error detected in Universal Action', { folderUid, requesterEmail });
      
      const fullAccessCard = cards.buildFolderApprovalCardWithStatus({
        approvalId,
        requesterName,
        requesterEmail,
        folderName,
        justification: data.justification || '',
        status: 'owner',
        statusMessage: 'User Already Has Full Access (Cannot Downgrade)',
        approverName: approver.name,
        permission,
        processedTime,
      });
      
      // Non-blocking update to avoid Teams timeout
      tryUpdateApprovalCard(approvalId, fullAccessCard, context)
        .then(() => log.debug('Updated channel card with full access status'))
        .catch(err => log.debug('Could not update channel card', err.message));
      
      return { updatedCard: fullAccessCard };
    }
    
    if (isPermissionConflictError(result.error)) {
      log.info('Permission conflict detected for folder in Universal Action', { folderUid, requesterEmail, error: result.error });
      return { error: result.error, keepCardActive: true };
    }
    
    if (isRecordOwnerError(result.error)) {
      log.info('Folder owner error detected in Universal Action', { folderUid, requesterEmail });
      
      const ownerStatusCard = cards.buildFolderApprovalCardWithStatus({
        approvalId,
        requesterName,
        requesterEmail,
        folderName,
        justification: data.justification || '',
        status: 'owner',
        statusMessage: 'User Already Has Full Access',
        approverName: approver.name,
        permission,
        processedTime,
      });
      
      // Non-blocking update to avoid Teams timeout
      tryUpdateApprovalCard(approvalId, ownerStatusCard, context)
        .then(() => log.debug('Updated channel card with owner status'))
        .catch(err => log.debug('Could not update channel card', err.message));
      
      return { updatedCard: ownerStatusCard };
    }
    
    log.error('Failed to grant folder access', result.error);
    return { error: result.error };
  }
  
  const expiresAtFormatted = formatExpiryDate(result.expiresAt);
  const isInvitationSent = result.invitationSent;
  
  storeApprovalStatus(approvalId, {
    status: isInvitationSent ? 'invitation_sent' : 'approved',
    type: 'folder',
    approverName: approver.name,
    requesterName,
    requesterEmail,
    folderName,
    justification: data.justification || '',
    permission,
    duration: getDisplayDuration(permission, duration, 'folder'),
    expiresAt: expiresAtFormatted,
    processedTime,
    invitationSent: isInvitationSent,
  });
  
  let updatedCard;
  
  if (isInvitationSent) {
    log.info('Share invitation sent for folder (user has no Keeper account)');
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
    updatedCard = cards.buildFolderApprovalCardWithStatus({
      approvalId,
      requesterName,
      requesterEmail,
      folderName,
      justification: data.justification || '',
      status: 'approved',
      approverName: approver.name,
      permission: permission,
      duration: getDisplayDuration(permission, duration, 'folder'),
      expiresAt: expiresAtFormatted,
      processedTime,
    });
  }
  
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
            duration: getDisplayDuration(permission, duration, 'folder'),
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
  
  // Non-blocking update to avoid Teams timeout
  tryUpdateApprovalCard(approvalId, updatedCard, context)
    .then(() => log.debug('Updated channel card with folder approved status'))
    .catch(err => log.debug('Could not update channel card', err.message));
  
  return { updatedCard };
}

async function handleRecordDenialWithCardResponse(context, data, approver) {
  const { approvalId, recordTitle, requesterName, requesterId, requesterEmail } = data;
  const processedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
  
  log.debug('Denying record via Universal Action', { approver: approver.name, recordTitle, requesterName });
  
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
  
  const updatedCard = cards.buildRecordApprovalCardWithStatus({
    approvalId,
    requesterName,
    recordTitle,
    justification: data.justification || '',
    status: 'denied',
    approverName: approver.name,
    processedTime,
  });
  
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
  
  // Non-blocking update to avoid Teams timeout
  tryUpdateApprovalCard(approvalId, updatedCard, context)
    .then(() => log.debug('Updated channel card with denied status'))
    .catch(err => log.debug('Could not update channel card', err.message));
  
  return { updatedCard };
}

async function handleFolderDenialWithCardResponse(context, data, approver) {
  const { approvalId, folderName, requesterName, requesterId, requesterEmail } = data;
  const processedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
  
  log.debug('Denying folder via Universal Action', { approver: approver.name, folderName, requesterName });
  
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
  
  const updatedCard = cards.buildFolderApprovalCardWithStatus({
    approvalId,
    requesterName,
    folderName,
    justification: data.justification || '',
    status: 'denied',
    approverName: approver.name,
    processedTime,
  });
  
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
  
  // Non-blocking update to avoid Teams timeout
  tryUpdateApprovalCard(approvalId, updatedCard, context)
    .then(() => log.debug('Updated channel card with folder denied status'))
    .catch(err => log.debug('Could not update channel card', err.message));
  
  return { updatedCard };
}

async function handleShareApprovalWithCardResponse(context, data, approver) {
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
  
  const result = await keeperClient.createOneTimeShare(
    recordUid,
    durationSeconds,
    editable
  );
  
  if (!result.success) {
    log.error('Failed to create one-time share', result.error);
    return { error: result.error };
  }
  
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
  
  if (requesterId) {
    try {
      const channelService = getChannelService();
      if (channelService) {
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
  
  // Non-blocking update to avoid Teams timeout
  tryUpdateApprovalCard(approvalId, updatedCard, context)
    .then(() => log.debug('Updated channel card with share approved status'))
    .catch(err => log.debug('Could not update channel card', err.message));
  
  return { updatedCard };
}

async function handleShareDenialWithCardResponse(context, data, approver) {
  const { approvalId, recordUid, recordTitle, requesterName, requesterId, requesterEmail } = data;
  const processedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
  
  log.debug('Denying one-time share via Universal Action', { approver: approver.name, recordTitle, requesterName });
  
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
  
  // Non-blocking update to avoid Teams timeout
  tryUpdateApprovalCard(approvalId, updatedCard, context)
    .then(() => log.debug('Updated channel card with share denied status'))
    .catch(err => log.debug('Could not update channel card', err.message));
  
  return { updatedCard };
}

module.exports = {
  routeApprovalAction,
  routeApprovalActionWithCardResponse,
  DURATION_MAP,
  parseDuration,
};
