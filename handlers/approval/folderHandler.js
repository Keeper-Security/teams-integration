/**
 * Folder Approval Handler
 * 
 * Handles approval and denial of folder access requests.
 */

const keeperClient = require('../../services/keeperClient');
const cards = require('../../cards');
const { getChannelService, createLogger } = require('../../services');
const { isPermissionConflictError, isRecordOwnerError } = require('../../utils/helpers');
const { 
  parseDuration, 
  getDisplayDuration,
  formatExpiryDate, 
  getApproverInfo, 
  tryUpdateApprovalCard,
  buildInvitationNotificationCard,
  buildPermissionConflictCard,
  getCurrentTimestamp,
} = require('./helpers');

const log = createLogger('FolderHandler');

/**
 * Handle approval of a folder access request
 */
async function handleFolderApproval(context, data) {
  log.info('Processing folder approval request', { approvalId: data.approvalId, folderUid: data.folderUid });
  
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
  
  const result = await keeperClient.grantFolderAccess(
    folderUid,
    requesterEmail,
    permission,
    durationSeconds
  );
  
  if (result.success) {
    const expiresAtFormatted = formatExpiryDate(result.expiresAt);
    const processedTime = getCurrentTimestamp();
    
    let updatedCard;
    
    if (result.invitationSent) {
      log.info('Share invitation sent to user for folder (no Keeper account)');
      
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
      updatedCard = cards.buildFolderApprovalCardWithStatus({
        approvalId: approvalId,
        requesterName: requesterName,
        requesterEmail: requesterEmail,
        folderName: folderName,
        justification: justification,
        status: 'approved',
        approverName: approver.name,
        permission: permission,
        duration: getDisplayDuration(permission, duration, 'folder'),
        expiresAt: expiresAtFormatted,
      });
    }
    
    // Non-blocking update to avoid Teams timeout
    tryUpdateApprovalCard(approvalId, updatedCard, context)
      .then(() => log.debug(`Successfully updated folder approval card with ${result.invitationSent ? 'INVITATION SENT' : 'APPROVED'} status`))
      .catch(err => {
        log.debug('Failed to update activity', err.message);
        context.send({
          type: 'message',
          attachments: [{
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: updatedCard,
          }],
        }).catch(sendErr => log.debug('Fallback send also failed', sendErr.message));
      });
    
    const requesterId = data.requesterId;
    if (requesterId) {
      try {
        const channelService = getChannelService();
        if (channelService) {
          let notificationCard;
          
          if (result.invitationSent) {
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
    if (result.alreadyHasAccess) {
      log.info('User already has access to folder', { folderUid, requesterEmail, currentPermission: result.currentPermission });
      
      const processedTime = getCurrentTimestamp();
      const alreadyHasAccessCard = cards.buildFolderAlreadyHasAccessCard({
        approvalId: approvalId,
        requesterName: requesterName,
        requesterEmail: requesterEmail,
        folderName: folderName,
        folderUid: folderUid,
        justification: justification,
        currentPermission: result.currentPermission,
        currentPermissionLabel: result.currentPermissionLabel,
        approverName: approver.name,
        processedTime: processedTime,
      });
      
      // Non-blocking update to avoid Teams timeout
      tryUpdateApprovalCard(approvalId, alreadyHasAccessCard, context)
        .then(() => log.debug('Updated approval card with already has access status'))
        .catch(err => {
          log.debug('Failed to update card', err.message);
          context.send({
            type: 'message',
            attachments: [{
              contentType: 'application/vnd.microsoft.card.adaptive',
              content: alreadyHasAccessCard,
            }],
          }).catch(sendErr => log.debug('Fallback send also failed', sendErr.message));
        });
    } else if (result.isFullAccessError) {
      log.info('Folder full access error detected', { folderUid, requesterEmail });
      
      try {
        const channelService = getChannelService();
        if (channelService) {
          const fullAccessErrorCard = {
            type: 'AdaptiveCard',
            '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
            version: '1.5',
            body: [
              { type: 'TextBlock', text: 'Permission Change Not Allowed', weight: 'Bolder', size: 'Large', color: 'Attention' },
              { type: 'TextBlock', text: `The selected user already has **Manage Users and Records** permission on this folder and cannot be downgraded.`, wrap: true, spacing: 'Medium' },
              { type: 'FactSet', spacing: 'Medium', facts: [
                { title: 'Request ID:', value: approvalId || 'N/A' },
                { title: 'Folder:', value: folderName },
                { title: 'Requester:', value: `${requesterName} (${requesterEmail})` },
              ]},
              { type: 'TextBlock', text: 'The user already has full access to this folder. No action is needed.', wrap: true, spacing: 'Medium', isSubtle: true },
            ],
          };
          
          await channelService.sendDirectMessage(approver.id, {
            type: 'message',
            attachments: [{ contentType: 'application/vnd.microsoft.card.adaptive', content: fullAccessErrorCard }],
          });
          log.debug('Sent full access error notification to approver');
        }
      } catch (notifyError) {
        log.error('Error sending full access notification to approver', notifyError.message);
      }
      
      const fullAccessCard = cards.buildFolderApprovalCardWithStatus({
        approvalId: approvalId,
        requesterName: requesterName,
        requesterEmail: requesterEmail,
        folderName: folderName,
        justification: justification,
        status: 'owner',
        statusMessage: 'User Already Has Full Access (Cannot Downgrade)',
        approverName: approver.name,
        permission: permission,
      });
      
      // Non-blocking update to avoid Teams timeout
      tryUpdateApprovalCard(approvalId, fullAccessCard, context)
        .then(() => log.debug('Updated approval card with full access status'))
        .catch(err => {
          log.debug('Failed to update card', err.message);
          context.send(`Cannot modify permissions: ${requesterEmail} already has "Manage Users and Records" permission and cannot be downgraded.`)
            .catch(sendErr => log.debug('Fallback send also failed', sendErr.message));
        });
    } else if (isPermissionConflictError(result.error)) {
      log.info('Permission conflict detected for folder approval', { folderUid, requesterEmail, error: result.error });
      
      try {
        const channelService = getChannelService();
        if (channelService) {
          const conflictCard = buildPermissionConflictCard({
            itemTitle: folderName,
            itemType: 'folder',
            requesterName: requesterName,
            requesterEmail: requesterEmail,
            requestedPermission: permission,
            errorMessage: result.error,
          });
          
          await channelService.sendDirectMessage(approver.id, {
            type: 'message',
            attachments: [{
              contentType: 'application/vnd.microsoft.card.adaptive',
              content: conflictCard,
            }],
          });
          log.debug('Sent permission conflict notification to approver');
        }
      } catch (notifyError) {
        log.error('Error sending conflict notification to approver', notifyError.message);
      }
    } else if (isRecordOwnerError(result.error)) {
      log.info('Folder owner error detected', { folderUid, requesterEmail });
      await context.send(`Cannot grant access: ${requesterEmail} is the owner of this folder.`);
    } else {
      await context.send('Failed to grant folder access: ' + result.error);
    }
  }
}

/**
 * Handle denial of a folder access request
 */
async function handleFolderDenial(context, data) {
  log.info('Processing folder denial request', { approvalId: data.approvalId, folderName: data.folderName });
  
  const approver = getApproverInfo(context.activity);
  const folderName = data.folderName || data.folderUid || 'Unknown Folder';
  const requesterName = data.requesterName || 'User';
  const approvalId = data.approvalId || 'N/A';
  const justification = data.justification || '';
  
  log.info('Denying folder access', { approver: approver.name, folderName, requesterName });
  
  const updatedCard = cards.buildFolderApprovalCardWithStatus({
    approvalId,
    requesterName,
    folderName,
    justification,
    status: 'denied',
    approverName: approver.name,
  });
  
  // Non-blocking update to avoid Teams timeout
  tryUpdateApprovalCard(approvalId, updatedCard, context)
    .then(() => log.debug('Updated original folder card with denied status'))
    .catch(err => {
      log.debug('Failed to update folder card', err.message);
      context.send({
        type: 'message',
        attachments: [{
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: updatedCard,
        }],
      }).catch(sendErr => log.debug('Fallback send also failed', sendErr.message));
    });
  
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
  
  log.info('Folder denial complete', { approvalId });
}

module.exports = {
  handleFolderApproval,
  handleFolderDenial,
};
