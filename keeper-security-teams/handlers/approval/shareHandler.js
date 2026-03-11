/**
 * Share Approval Handler
 * 
 * Handles approval and denial of one-time share requests.
 */

const keeperClient = require('../../services/keeperClient');
const cards = require('../../cards');
const { getChannelService, createLogger } = require('../../services');
const { isPermissionConflictError, isPamRecordError } = require('../../utils/helpers');
const { 
  parseDuration, 
  getApproverInfo, 
  buildPermissionConflictCard,
} = require('./helpers');

const log = createLogger('ShareHandler');

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
    if (isPermissionConflictError(result.error)) {
      log.info('Permission conflict detected for one-time share', { recordUid, error: result.error });
      
      try {
        const channelService = getChannelService();
        if (channelService) {
          const conflictCard = buildPermissionConflictCard({
            itemTitle: recordTitle,
            itemType: 'record',
            requesterName: requesterName,
            requesterEmail: 'N/A (One-Time Share)',
            requestedPermission: editable ? 'Editable Share' : 'View-Only Share',
            errorMessage: result.error,
          });
          
          await channelService.sendDirectMessage(approver.id, {
            type: 'message',
            attachments: [{
              contentType: 'application/vnd.microsoft.card.adaptive',
              content: conflictCard,
            }],
          });
          log.debug('Sent permission conflict notification to approver for one-time share');
        }
      } catch (notifyError) {
        log.error('Error sending conflict notification to approver', notifyError.message);
      }
    } else if (isPamRecordError(result.error)) {
      log.info('PAM record error detected for one-time share', { recordUid, error: result.error });
      await context.send(`**One-Time Share Not Available**\n\nThe record \`${recordTitle}\` is a PAM record.\n\nOne-Time Shares are currently not available for PAM records. The requester should use \`keeper-request-record\` to request direct access instead.`);
    } else {
      await context.send('Failed to create share link: ' + result.error);
    }
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

module.exports = {
  handleShareApproval,
  handleShareDenial,
};
