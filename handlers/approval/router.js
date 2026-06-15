/**
 * Approval Router
 * 
 * Routes card actions to appropriate handlers.
 */

const keeperClient = require('../../services/keeperClient');
const cards = require('../../cards');
const { getChannelService, getApprovalStatus, storeApprovalStatus, createLogger } = require('../../services');
const { isPermissionConflictError, isRecordOwnerError, isPamUserRecordType } = require('../../utils/helpers');
const { 
  DURATION_MAP,
  parseDuration, 
  getDisplayDuration,
  formatExpiryDate, 
  getApproverInfo, 
  tryUpdateApprovalCard,
  buildInvitationNotificationCard,
  getCurrentTimestamp,
  sanitizeDisplayField,
  isValidEmail,
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
        version: '1.5',
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
  const { approvalId, recordUid, requesterId, requesterEmail } = data;
  const recordTitle   = sanitizeDisplayField(data.recordTitle);
  const requesterName = sanitizeDisplayField(data.requesterName);
  const permission = data.permission || 'view_only';
  const duration = data.duration || '24h';
  const durationSeconds = parseDuration(duration);
  const rotateOnExpire = data.rotateOnExpire === 'true' || data.rotateOnExpire === true;
  const isNsf = data.isNsf === 'true' || data.isNsf === true;
  const nsfRole = data.nsfRole || 'viewer';
  const processedTime = getCurrentTimestamp();
  
  log.info('Approving record via Universal Action', { approvalId, approver: approver.name, recordTitle, permission, duration, rotateOnExpire, isNsf, nsfRole });
  
  if (!recordUid) {
    log.error('Missing record UID');
    return { error: 'Missing record UID' };
  }
  
  if (!requesterEmail || !isValidEmail(requesterEmail)) {
    log.error('Missing or invalid requester email');
    return { error: 'Missing or invalid requester email' };
  }
  
  // Build and return the "Processing" card immediately to avoid Teams timeout
  const processingCard = cards.buildRecordProcessingCard({
    approvalId,
    requesterName,
    requesterEmail,
    recordTitle,
    justification: data.justification || '',
    permission: isNsf ? nsfRole : permission,
    duration: getDisplayDuration(isNsf ? nsfRole : permission, duration, 'record'),
    approverName: approver.name,
    processedTime,
  });
  
  // Fire off the grant operation asynchronously - don't await
  processRecordGrantAsync(context, data, approver, {
    approvalId,
    recordUid,
    recordTitle,
    requesterName,
    requesterId,
    requesterEmail,
    permission,
    duration,
    durationSeconds,
    rotateOnExpire,
    isNsf,
    nsfRole,
    processedTime,
  });
  
  // Return the processing card immediately
  return { updatedCard: processingCard };
}

/**
 * Process the record grant asynchronously after returning the processing card
 */
async function processRecordGrantAsync(context, data, approver, params) {
  const {
    approvalId,
    recordUid,
    recordTitle,
    requesterName,
    requesterId,
    requesterEmail,
    permission,
    duration,
    durationSeconds,
    rotateOnExpire,
    isNsf = false,
    nsfRole = 'viewer',
    processedTime,
  } = params;

  // For NSF grants the "permission" shown/stored is the NSF role.
  const displayPermission = isNsf ? nsfRole : permission;
  
  try {
    let result;

    if (isNsf) {
      // Nested Share Folder record grant uses a single role and supports
      // a 'permanent' (never-expiring) duration when no expiry is selected.
      // Transfer Ownership ('owner') is always permanent regardless of the
      // selected duration.
      const nsfDurationSeconds = (nsfRole === 'owner' || duration === 'permanent') ? null : durationSeconds;
      result = await keeperClient.grantNsfRecordAccess(
        recordUid,
        requesterEmail,
        nsfRole,
        nsfDurationSeconds
      );
    } else {
      // Centralized server-side re-check: only honor rotateOnExpire when the
      // target record is actually a PAM User record. This gates both the direct
      // approve_record path and the multi-select approve_selected_record path,
      // so a tampered/replayed payload can't force rotation on an ineligible record.
      let effectiveRotateOnExpire = rotateOnExpire;
      if (effectiveRotateOnExpire) {
        let recordType = '';
        try {
          const rec = await keeperClient.getRecordByUid(recordUid);
          recordType = rec?.recordType || '';
        } catch (e) {
          log.debug('Record type lookup for rotate-on-expire gate failed, treating as ineligible', e.message);
        }
        if (!isPamUserRecordType(recordType)) {
          log.warn('rotateOnExpire requested for non-PAM record, ignoring flag', { recordUid, recordType });
          effectiveRotateOnExpire = false;
        }
      }

      result = await keeperClient.grantRecordAccess(
        recordUid,
        requesterEmail,
        permission,
        durationSeconds,
        effectiveRotateOnExpire
      );
    }
    
    let finalCard;
    
    if (!result.success) {
      if (result.alreadyHasAccess) {
        log.info('User already has access to record', { recordUid, requesterEmail, currentPermission: result.currentPermission });
        
        finalCard = cards.buildRecordAlreadyHasAccessCard({
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
      } else if (result.isOwnerError || isRecordOwnerError(result.error)) {
        log.info('Record owner error detected', { recordUid, requesterEmail });
        
        finalCard = cards.buildRecordApprovalCardWithStatus({
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
      } else if (result.errorCode === 'pam_rotation_not_configured') {
        log.warn('PAM rotation not configured, re-rendering approval card with error banner', { recordUid });

        finalCard = cards.buildRecordApprovalCard({
          approvalId,
          requesterName,
          requesterId,
          requesterEmail,
          recordTitle,
          recordUid,
          recordType: data.recordType || '',
          justification: data.justification || '',
          isUid: true,
          identifier: recordTitle,
          errorBanner: 'Rotation is not configured on this record. Disable "Rotate credentials when access expires" and approve again.',
        });
      } else {
        log.error('Failed to grant record access', { error: result.error });
        
        // Build an error card
        finalCard = cards.buildRecordApprovalCardWithStatus({
          approvalId,
          requesterName,
          requesterEmail,
          recordTitle,
          justification: data.justification || '',
          status: 'denied',
          statusMessage: `Failed: ${result.error || 'Unknown error'}`,
          approverName: approver.name,
          permission,
          processedTime,
        });
      }
    } else {
      // Success!
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
        permission: displayPermission,
        duration: getDisplayDuration(displayPermission, duration, 'record'),
        expiresAt: expiresAtFormatted,
        processedTime,
        invitationSent: isInvitationSent,
        isNsf,
      });
      
      const pamRotateScheduled = !!result.rotateOnExpire;

      if (isInvitationSent) {
        log.info('Share invitation sent for record (user has no Keeper account)');
        finalCard = cards.buildRecordInvitationSentCard({
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
        finalCard = cards.buildRecordApprovalCardWithStatus({
          approvalId,
          requesterName,
          requesterEmail,
          recordTitle,
          justification: data.justification || '',
          status: 'approved',
          approverName: approver.name,
          permission: displayPermission,
          duration: getDisplayDuration(displayPermission, duration, 'record'),
          expiresAt: expiresAtFormatted,
          processedTime,
          rotateOnExpire: pamRotateScheduled,
          isNsf,
        });
      }
      
      // Send notification to requester
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
                permission: displayPermission,
                duration: getDisplayDuration(displayPermission, duration, 'record'),
                expiresAt: expiresAtFormatted,
                approverName: approver.name,
                itemType: 'record',
                rotateOnExpire: pamRotateScheduled,
                isNsf,
              });
            }
            
            channelService.sendDirectMessage(requesterId, {
              type: 'message',
              attachments: [{
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: notificationCard,
              }],
            }).catch(err => {
              log.debug('Could not send notification to requester', err.message);
            });
          }
        } catch (notifyError) {
          log.error('Error sending notification', notifyError.message);
        }
      }
    }
    
    // Update the card with the final status
    if (finalCard) {
      try {
        await tryUpdateApprovalCard(approvalId, finalCard, context);
        log.info('Updated record approval card with final status', { approvalId });
      } catch (updateErr) {
        log.error('Failed to update record approval card', { approvalId, error: updateErr.message });
      }
    }
  } catch (error) {
    log.error('Error in async record grant processing', { approvalId, error: error.message });
    
    // Try to update the card with an error status
    try {
      const errorCard = cards.buildRecordApprovalCardWithStatus({
        approvalId,
        requesterName,
        requesterEmail,
        recordTitle,
        justification: data.justification || '',
        status: 'denied',
        statusMessage: `Error: ${error.message}`,
        approverName: approver.name,
        permission,
        processedTime,
      });
      
      await tryUpdateApprovalCard(approvalId, errorCard, context);
    } catch (updateErr) {
      log.error('Failed to update card with error status', { approvalId, error: updateErr.message });
    }
  }
}

async function handleFolderApprovalWithCardResponse(context, data, approver) {
  const { approvalId, folderUid, requesterId, requesterEmail } = data;
  const folderName    = sanitizeDisplayField(data.folderName);
  const requesterName = sanitizeDisplayField(data.requesterName);
  const permission = data.permission || 'no_permissions';
  const duration = data.duration || '24h';
  const durationSeconds = parseDuration(duration);
  const rotateOnExpire = data.rotateOnExpire === 'true' || data.rotateOnExpire === true;
  const isNsf = data.isNsf === 'true' || data.isNsf === true;
  const nsfRole = data.nsfRole || 'viewer';
  const processedTime = getCurrentTimestamp();
  
  log.info('Approving folder via Universal Action', { approvalId, approver: approver.name, folderName, permission, duration, rotateOnExpire, isNsf, nsfRole });
  
  if (!folderUid) {
    log.error('Missing folder UID');
    return { error: 'Missing folder UID' };
  }
  
  if (!requesterEmail || !isValidEmail(requesterEmail)) {
    log.error('Missing or invalid requester email');
    return { error: 'Missing or invalid requester email' };
  }
  
  // Build and return the "Processing" card immediately to avoid Teams timeout
  const processingCard = cards.buildFolderProcessingCard({
    approvalId,
    requesterName,
    requesterEmail,
    folderName,
    justification: data.justification || '',
    permission: isNsf ? nsfRole : permission,
    duration: getDisplayDuration(isNsf ? nsfRole : permission, duration, 'folder'),
    approverName: approver.name,
    processedTime,
  });
  
  // Fire off the grant operation asynchronously - don't await
  processFolderGrantAsync(context, data, approver, {
    approvalId,
    folderUid,
    folderName,
    requesterName,
    requesterId,
    requesterEmail,
    permission,
    duration,
    durationSeconds,
    rotateOnExpire,
    isNsf,
    nsfRole,
    processedTime,
  });
  
  // Return the processing card immediately
  return { updatedCard: processingCard };
}

/**
 * Process the folder grant asynchronously after returning the processing card
 */
async function processFolderGrantAsync(context, data, approver, params) {
  const {
    approvalId,
    folderUid,
    folderName,
    requesterName,
    requesterId,
    requesterEmail,
    permission,
    duration,
    durationSeconds,
    rotateOnExpire,
    isNsf = false,
    nsfRole = 'viewer',
    processedTime,
  } = params;

  // For NSF grants the "permission" shown/stored is the NSF role.
  const displayPermission = isNsf ? nsfRole : permission;
  
  try {
    let result;

    if (isNsf) {
      // Nested Share Folder grant uses a single role and supports a
      // 'permanent' (never-expiring) duration when no expiry is selected.
      const nsfDurationSeconds = duration === 'permanent' ? null : durationSeconds;
      result = await keeperClient.grantNsfFolderAccess(
        folderUid,
        requesterEmail,
        nsfRole,
        nsfDurationSeconds
      );
    } else {
      // Centralized server-side re-check: only honor rotateOnExpire when the
      // target folder is actually ROE-eligible. This gates both the direct
      // approve_folder path and the multi-select approve_selected_folder path,
      // so a tampered/replayed payload can't force rotation on an ineligible folder.
      let effectiveRotateOnExpire = rotateOnExpire;
      if (effectiveRotateOnExpire) {
        let eligible = false;
        try {
          eligible = await keeperClient.isPamUserFolder(folderUid);
        } catch (e) {
          log.debug('Folder ROE-eligibility lookup for rotate-on-expire gate failed, treating as ineligible', e.message);
        }
        if (!eligible) {
          log.warn('rotateOnExpire requested for non-eligible folder, ignoring flag', { folderUid });
          effectiveRotateOnExpire = false;
        }
      }

      result = await keeperClient.grantFolderAccess(
        folderUid,
        requesterEmail,
        permission,
        durationSeconds,
        effectiveRotateOnExpire
      );
    }
    
    let finalCard;
    
    if (!result.success) {
      if (result.alreadyHasAccess) {
        log.info('User already has access to folder', { folderUid, requesterEmail, currentPermission: result.currentPermission });
        
        finalCard = cards.buildFolderAlreadyHasAccessCard({
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
      } else if (result.isFullAccessError) {
        log.info('Folder full access error detected', { folderUid, requesterEmail });
        
        finalCard = cards.buildFolderApprovalCardWithStatus({
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
      } else if (isRecordOwnerError(result.error)) {
        log.info('Folder owner error detected', { folderUid, requesterEmail });
        
        finalCard = cards.buildFolderApprovalCardWithStatus({
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
      } else if (result.errorCode === 'pam_rotation_not_configured') {
        log.warn('PAM rotation not configured for folder, re-rendering approval card with error banner', { folderUid });

        finalCard = cards.buildFolderApprovalCard({
          approvalId,
          requesterName,
          requesterId,
          requesterEmail,
          folderName,
          folderUid,
          justification: data.justification || '',
          isUid: true,
          identifier: folderName,
          isPamFolder: true,
          errorBanner: 'Rotation is not configured on this folder\'s records. Disable "Rotate credentials when access expires" and approve again.',
        });
      } else {
        log.error('Failed to grant folder access', { error: result.error });
        
        // Build an error card
        finalCard = cards.buildFolderApprovalCardWithStatus({
          approvalId,
          requesterName,
          requesterEmail,
          folderName,
          justification: data.justification || '',
          status: 'denied',
          statusMessage: `Failed: ${result.error || 'Unknown error'}`,
          approverName: approver.name,
          permission,
          processedTime,
        });
      }
    } else {
      // Success!
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
        permission: displayPermission,
        duration: getDisplayDuration(displayPermission, duration, 'folder'),
        expiresAt: expiresAtFormatted,
        processedTime,
        invitationSent: isInvitationSent,
        isNsf,
      });
      
      const pamRotateScheduled = !!result.rotateOnExpire;

      if (isInvitationSent) {
        log.info('Share invitation sent for folder (user has no Keeper account)');
        finalCard = cards.buildFolderInvitationSentCard({
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
        finalCard = cards.buildFolderApprovalCardWithStatus({
          approvalId,
          requesterName,
          requesterEmail,
          folderName,
          justification: data.justification || '',
          status: 'approved',
          approverName: approver.name,
          permission: displayPermission,
          duration: getDisplayDuration(displayPermission, duration, 'folder'),
          expiresAt: expiresAtFormatted,
          processedTime,
          rotateOnExpire: pamRotateScheduled,
          isNsf,
        });
      }
      
      // Send notification to requester
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
                permission: displayPermission,
                duration: getDisplayDuration(displayPermission, duration, 'folder'),
                expiresAt: expiresAtFormatted,
                approverName: approver.name,
                rotateOnExpire: pamRotateScheduled,
                isNsf,
              });
            }
            
            channelService.sendDirectMessage(requesterId, {
              type: 'message',
              attachments: [{
                contentType: 'application/vnd.microsoft.card.adaptive',
                content: notificationCard,
              }],
            }).catch(err => {
              log.debug('Could not send notification to requester', err.message);
            });
          }
        } catch (notifyError) {
          log.error('Error sending notification', notifyError.message);
        }
      }
    }
    
    // Update the card with the final status
    if (finalCard) {
      try {
        await tryUpdateApprovalCard(approvalId, finalCard, context);
        log.info('Updated folder approval card with final status', { approvalId });
      } catch (updateErr) {
        log.error('Failed to update folder approval card', { approvalId, error: updateErr.message });
      }
    }
  } catch (error) {
    log.error('Error in async folder grant processing', { approvalId, error: error.message });
    
    // Try to update the card with an error status
    try {
      const errorCard = cards.buildFolderApprovalCardWithStatus({
        approvalId,
        requesterName,
        requesterEmail,
        folderName,
        justification: data.justification || '',
        status: 'denied',
        statusMessage: `Error: ${error.message}`,
        approverName: approver.name,
        permission,
        processedTime,
      });
      
      await tryUpdateApprovalCard(approvalId, errorCard, context);
    } catch (updateErr) {
      log.error('Failed to update card with error status', { approvalId, error: updateErr.message });
    }
  }
}

async function handleRecordDenialWithCardResponse(context, data, approver) {
  const { approvalId, recordTitle, requesterName, requesterId, requesterEmail } = data;
  const processedTime = getCurrentTimestamp();
  
  log.info('Denying record via Universal Action', { approvalId, approver: approver.name, recordTitle });
  
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
  const processedTime = getCurrentTimestamp();
  
  log.info('Denying folder via Universal Action', { approvalId, approver: approver.name, folderName });
  
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
  const { approvalId, recordUid, requesterId, requesterEmail } = data;
  const recordTitle   = sanitizeDisplayField(data.recordTitle);
  const requesterName = sanitizeDisplayField(data.requesterName);
  const duration = data.duration || '24h';
  const durationSeconds = parseDuration(duration) || 86400;
  const editable = data.editable === 'true' || data.editable === true;
  const processedTime = getCurrentTimestamp();
  
  log.info('Approving one-time share via Universal Action', { approvalId, approver: approver.name, recordTitle, duration, editable });
  
  if (!recordUid) {
    log.error('Missing record UID for share');
    return { error: 'Missing record UID' };
  }

  if (requesterEmail && !isValidEmail(requesterEmail)) {
    log.warn('Invalid requesterEmail format rejected in share approval', { approvalId });
    return { error: 'Invalid requester email format' };
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
  
  // formatExpiryDate renders via Adaptive Card DATE()/TIME() so the recipient
  // sees the expiry in their own local timezone (consistent with "Processed").
  const expiresAtFormatted = result.expiresAt ? formatExpiryDate(result.expiresAt) : 'N/A';
  
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
  const processedTime = getCurrentTimestamp();
  
  log.info('Denying one-time share via Universal Action', { approvalId, approver: approver.name, recordTitle });
  
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
