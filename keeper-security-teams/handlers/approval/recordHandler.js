/**
 * Record Approval Handler
 * 
 * Handles approval and denial of record access requests.
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
} = require('./helpers');

const log = createLogger('RecordHandler');

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
  
  const result = await keeperClient.grantRecordAccess(
    recordUid,
    requesterEmail,
    permission,
    durationSeconds
  );
  
  if (result.success) {
    const expiresAtFormatted = formatExpiryDate(result.expiresAt);
    const processedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
    
    let updatedCard;
    
    if (result.invitationSent) {
      log.info('Share invitation sent to user (no Keeper account)');
      
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
      updatedCard = cards.buildRecordApprovalCardWithStatus({
        approvalId: approvalId,
        requesterName: requesterName,
        requesterEmail: requesterEmail,
        recordTitle: recordTitle,
        justification: justification,
        status: 'approved',
        approverName: approver.name,
        permission: permission,
        duration: getDisplayDuration(permission, duration, 'record'),
        expiresAt: expiresAtFormatted,
      });
    }
    
    // Non-blocking update to avoid Teams timeout
    tryUpdateApprovalCard(approvalId, updatedCard, context)
      .then(() => log.debug('Successfully updated approval card with', result.invitationSent ? 'INVITATION SENT' : 'APPROVED', 'status'))
      .catch(err => {
        log.debug('Failed to update activity', err.message);
        // Fire-and-forget fallback
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
    if (result.alreadyHasAccess) {
      log.info('User already has access to record', { recordUid, requesterEmail, currentPermission: result.currentPermission });
      
      const processedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
      const alreadyHasAccessCard = cards.buildRecordAlreadyHasAccessCard({
        approvalId: approvalId,
        requesterName: requesterName,
        requesterEmail: requesterEmail,
        recordTitle: recordTitle,
        recordUid: recordUid,
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
    } else if (isPermissionConflictError(result.error)) {
      log.info('Permission conflict detected for record approval', { recordUid, requesterEmail, error: result.error });
      
      try {
        const channelService = getChannelService();
        if (channelService) {
          const conflictCard = buildPermissionConflictCard({
            itemTitle: recordTitle,
            itemType: 'record',
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
    } else if (isRecordOwnerError(result.error) || result.isOwnerError) {
      log.info('Record owner error detected', { recordUid, requesterEmail });
      
      try {
        const channelService = getChannelService();
        if (channelService) {
          const ownerErrorCard = {
            type: 'AdaptiveCard',
            '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
            version: '1.4',
            body: [
              { type: 'TextBlock', text: 'Access Grant Failed', weight: 'Bolder', size: 'Large', color: 'Attention' },
              { type: 'TextBlock', text: `The selected user is the current **owner** of this record and already has full permissions.`, wrap: true, spacing: 'Medium' },
              { type: 'FactSet', spacing: 'Medium', facts: [
                { title: 'Request ID:', value: approvalId || 'N/A' },
                { title: 'Record:', value: recordTitle },
                { title: 'Requester:', value: `${requesterName} (${requesterEmail})` },
              ]},
              { type: 'TextBlock', text: 'No action is needed - the user already has full access to this record.', wrap: true, spacing: 'Medium', isSubtle: true },
            ],
          };
          
          await channelService.sendDirectMessage(approver.id, {
            type: 'message',
            attachments: [{ contentType: 'application/vnd.microsoft.card.adaptive', content: ownerErrorCard }],
          });
          log.debug('Sent owner error notification to approver');
        }
      } catch (notifyError) {
        log.error('Error sending owner notification to approver', notifyError.message);
      }
      
      const processedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
      const ownerStatusCard = cards.buildRecordApprovalCardWithStatus({
        approvalId: approvalId,
        requesterName: requesterName,
        requesterEmail: requesterEmail,
        recordTitle: recordTitle,
        justification: justification,
        status: 'owner',
        statusMessage: 'User Already Has Full Access (Owner)',
        approverName: approver.name,
        permission: permission,
      });
      
      // Non-blocking update to avoid Teams timeout
      tryUpdateApprovalCard(approvalId, ownerStatusCard, context)
        .then(() => log.debug('Updated approval card with owner status'))
        .catch(err => {
          log.debug('Failed to update card', err.message);
          context.send(`Cannot modify permissions: ${requesterEmail} is the owner of this record and already has full access.`)
            .catch(sendErr => log.debug('Fallback send also failed', sendErr.message));
        });
    } else {
      context.send('Failed to grant access: ' + result.error)
        .catch(err => log.debug('Failed to send error message', err.message));
    }
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
  
  const updatedCard = cards.buildRecordApprovalCardWithStatus({
    approvalId,
    requesterName,
    recordTitle,
    justification,
    status: 'denied',
    approverName: approver.name,
  });
  
  // Non-blocking update to avoid Teams timeout
  tryUpdateApprovalCard(approvalId, updatedCard, context)
    .then(() => log.debug('Updated original card with denied status'))
    .catch(err => {
      log.debug('Failed to update card', err.message);
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

module.exports = {
  handleRecordApproval,
  handleRecordDenial,
};
