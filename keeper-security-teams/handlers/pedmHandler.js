/**
 * PEDM (Privileged Elevation & Delegation Management) Handler
 * 
 * Handles Adaptive Card action submissions for PEDM elevation requests:
 * - Approve PEDM request
 * - Deny PEDM request
 * 
 * Features:
 * - In-place card updates (not new messages)
 * - Already processed request handling
 * - Channel card updates for all users via message edit
 */

const keeperClient = require('../services/keeperClient');
const cards = require('../cards');
const { createLogger, getChannelService } = require('../services');

const log = createLogger('PedmHandler');

/**
 * Helper function to update a PEDM approval card in the channel
 * Uses context.updateActivity() first (most reliable in invoke handlers),
 * then falls back to channelService.updateApprovalCard()
 * @param {string} approvalUid - The PEDM approval UID
 * @param {Object} updatedCard - The updated Adaptive Card content
 * @param {Object} context - The Teams context
 */
async function tryUpdatePedmCard(approvalUid, updatedCard, context) {
  // replyToId is the activity ID of the card being interacted with
  const activityId = context?.activity?.replyToId;
  
  if (!activityId) {
    log.debug(`No replyToId found for PEDM ${approvalUid}`);
    return false;
  }

  log.debug(`Updating PEDM card ${approvalUid}`, { activityId });

  // Method 1: Try context.updateActivity() directly (most reliable in invoke handlers)
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
      log.info('Updated PEDM card via context.updateActivity()', { approvalUid, activityId });
      return true;
    } catch (contextError) {
      log.debug('context.updateActivity() failed, trying fallback', contextError.message);
    }
  }

  // Method 2: Fallback to channel service
  const channelService = getChannelService();
  if (!channelService) {
    log.debug('Channel service not available');
    return false;
  }

  const success = await channelService.updateApprovalCard(activityId, updatedCard);
  if (success) {
    log.info('Updated PEDM card via channelService', { approvalUid, activityId });
    return true;
  }

  log.debug('Failed to update PEDM card in channel');
  return false;
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
 * Build an "already processed" card when the request was handled elsewhere
 */
function buildAlreadyProcessedCard(username, approvalUid) {
  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: [
      {
        type: 'Container',
        style: 'warning',
        items: [
          {
            type: 'TextBlock',
            text: 'Request Already Processed',
            weight: 'Bolder',
            size: 'Medium',
          },
        ],
      },
      {
        type: 'TextBlock',
        text: 'This EPM request for **' + username + '** has already been processed by another admin or has expired.',
        wrap: true,
        spacing: 'Medium',
      },
      {
        type: 'FactSet',
        facts: [
          { title: 'Approval UID', value: approvalUid },
          { title: 'Status', value: 'Already Processed' },
        ],
        spacing: 'Medium',
      },
    ],
  };
}

/**
 * Handle approval of a PEDM elevation request
 * Returns Adaptive Card for in-place update and updates channel card for all users
 */
async function handlePedmApproval(context, data) {
  const approver = getApproverInfo(context.activity);
  const approvalUid = data.approvalUid;
  const username = data.username || 'Unknown User';
  const command = data.command || 'Unknown Command';
  const agentUid = data.agentUid || '';
  
  if (!approvalUid) {
    // Return error card for in-place update
    return {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.4',
      body: [{
        type: 'TextBlock',
        text: 'Error: Missing approval UID',
        color: 'Attention',
      }],
    };
  }
  
  log.debug('Approving request', approvalUid);
  
  const result = await keeperClient.approvePedmRequest(approvalUid);
  
  let updatedCard;
  
  if (result.success) {
    log.info('PEDM request approved', { approvalUid, approver: approver.name });
    updatedCard = cards.buildPedmApprovedCard(
      approver.name,
      username,
      command,
      approvalUid,
      agentUid
    );
  } else if (result.already_processed) {
    log.debug('Request already processed', approvalUid);
    updatedCard = buildAlreadyProcessedCard(username, approvalUid);
  } else {
    // Return error card
    updatedCard = {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: 'Failed to Approve EPM Request',
          weight: 'Bolder',
          color: 'Attention',
        },
        {
          type: 'TextBlock',
          text: result.error || 'Unknown error',
          wrap: true,
        },
      ],
    };
  }
  
  // Update the channel card for all users
  try {
    await tryUpdatePedmCard(approvalUid, updatedCard, context);
  } catch (updateError) {
    log.debug('Could not update channel card', updateError.message);
  }
  
  return updatedCard;
}

/**
 * Handle denial of a PEDM elevation request
 * Returns Adaptive Card for in-place update and updates channel card for all users
 */
async function handlePedmDenial(context, data) {
  const approver = getApproverInfo(context.activity);
  const approvalUid = data.approvalUid;
  const username = data.username || 'Unknown User';
  const command = data.command || 'Unknown Command';
  const agentUid = data.agentUid || '';
  
  if (!approvalUid) {
    // Return error card for in-place update
    return {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.4',
      body: [{
        type: 'TextBlock',
        text: 'Error: Missing approval UID',
        color: 'Attention',
      }],
    };
  }
  
  log.debug('Denying request', approvalUid);
  
  const result = await keeperClient.denyPedmRequest(approvalUid);
  
  let updatedCard;
  
  if (result.success) {
    log.info('PEDM request denied', { approvalUid, approver: approver.name });
    updatedCard = cards.buildPedmDeniedCard(
      approver.name,
      username,
      command,
      approvalUid,
      agentUid
    );
  } else if (result.already_processed) {
    log.debug('Request already processed', approvalUid);
    updatedCard = buildAlreadyProcessedCard(username, approvalUid);
  } else {
    // Return error card
    updatedCard = {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: 'Failed to Deny EPM Request',
          weight: 'Bolder',
          color: 'Attention',
        },
        {
          type: 'TextBlock',
          text: result.error || 'Unknown error',
          wrap: true,
        },
      ],
    };
  }
  
  // Update the channel card for all users
  try {
    await tryUpdatePedmCard(approvalUid, updatedCard, context);
  } catch (updateError) {
    log.debug('Could not update channel card', updateError.message);
  }
  
  return updatedCard;
}

/**
 * Route PEDM card action to appropriate handler
 * Returns an Adaptive Card for in-place update (or null if action not handled)
 */
async function routePedmAction(context, data) {
  const action = data.action;
  
  switch (action) {
    case 'approve_pedm':
      return await handlePedmApproval(context, data);
      
    case 'deny_pedm':
      return await handlePedmDenial(context, data);
      
    default:
      return null;
  }
}

module.exports = {
  routePedmAction,
  handlePedmApproval,
  handlePedmDenial,
};
