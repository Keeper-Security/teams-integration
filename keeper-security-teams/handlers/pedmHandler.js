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
 */

const keeperClient = require('../services/keeperClient');
const cards = require('../cards');
const { createLogger } = require('../services');

const log = createLogger('PedmHandler');

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
 * Returns Adaptive Card for in-place update
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
  
  if (result.success) {
    // Return approved card for in-place update
    return cards.buildPedmApprovedCard(
      approver.name,
      username,
      command,
      approvalUid,
      agentUid
    );
  } else if (result.already_processed) {
    // Return "already processed" card
    log.debug('Request already processed', approvalUid);
    return buildAlreadyProcessedCard(username, approvalUid);
  } else {
    // Return error card
    return {
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
}

/**
 * Handle denial of a PEDM elevation request
 * Returns Adaptive Card for in-place update
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
  
  if (result.success) {
    // Return denied card for in-place update
    return cards.buildPedmDeniedCard(
      approver.name,
      username,
      command,
      approvalUid,
      agentUid
    );
  } else if (result.already_processed) {
    // Return "already processed" card
    log.debug('Request already processed', approvalUid);
    return buildAlreadyProcessedCard(username, approvalUid);
  } else {
    // Return error card
    return {
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
