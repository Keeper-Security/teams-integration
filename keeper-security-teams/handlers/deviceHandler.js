/**
 * Device Approval Handler
 * 
 * Handles Adaptive Card action submissions for Cloud SSO device approvals:
 * - Approve device
 * - Deny device
 * 
 * Features (matching Slack implementation):
 * - In-place card updates (not new messages)
 * - Already processed request handling
 */

const keeperClient = require('../services/keeperClient');
const cards = require('../cards');

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
 * Build an "already processed" card when the device request was handled elsewhere
 */
function buildAlreadyProcessedCard(username, deviceId) {
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
        text: 'This device approval request for **' + username + '** has already been processed by another admin or has expired.',
        wrap: true,
        spacing: 'Medium',
      },
      {
        type: 'FactSet',
        facts: [
          { title: 'Device ID', value: deviceId },
          { title: 'Status', value: 'Already Processed' },
        ],
        spacing: 'Medium',
      },
    ],
  };
}

/**
 * Handle approval of a device request
 * Returns Adaptive Card for in-place update
 */
async function handleDeviceApproval(context, data) {
  const approver = getApproverInfo(context.activity);
  const deviceId = data.deviceId;
  const deviceName = data.deviceName || 'Unknown Device';
  const username = data.username || 'Unknown User';
  
  if (!deviceId) {
    // Return error card for in-place update
    return {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.4',
      body: [{
        type: 'TextBlock',
        text: 'Error: Missing device ID',
        color: 'Attention',
      }],
    };
  }
  
  console.log('[Device Handler] Approving device:', deviceId);
  
  const result = await keeperClient.approveDevice(deviceId);
  
  if (result.success) {
    // Return approved card for in-place update
    return cards.buildDeviceApprovedCard(
      approver.name,
      deviceName,
      username,
      deviceId
    );
  } else if (result.already_processed) {
    // Return "already processed" card (like Slack does)
    console.log('[Device Handler] Device already processed:', deviceId);
    return buildAlreadyProcessedCard(username, deviceId);
  } else {
    // Return error card
    return {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: 'Failed to Approve Device',
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
 * Handle denial of a device request
 * Returns Adaptive Card for in-place update
 */
async function handleDeviceDenial(context, data) {
  const approver = getApproverInfo(context.activity);
  const deviceId = data.deviceId;
  const deviceName = data.deviceName || 'Unknown Device';
  const username = data.username || 'Unknown User';
  
  if (!deviceId) {
    // Return error card for in-place update
    return {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.4',
      body: [{
        type: 'TextBlock',
        text: 'Error: Missing device ID',
        color: 'Attention',
      }],
    };
  }
  
  console.log('[Device Handler] Denying device:', deviceId);
  
  const result = await keeperClient.denyDevice(deviceId);
  
  if (result.success) {
    // Return denied card for in-place update
    return cards.buildDeviceDeniedCard(
      approver.name,
      deviceName,
      username,
      deviceId
    );
  } else if (result.already_processed) {
    // Return "already processed" card (like Slack does)
    console.log('[Device Handler] Device already processed:', deviceId);
    return buildAlreadyProcessedCard(username, deviceId);
  } else {
    // Return error card
    return {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.4',
      body: [
        {
          type: 'TextBlock',
          text: 'Failed to Deny Device',
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
 * Route device card action to appropriate handler
 * Returns an Adaptive Card for in-place update (or null if action not handled)
 */
async function routeDeviceAction(context, data) {
  const action = data.action;
  
  switch (action) {
    case 'approve_device':
      return await handleDeviceApproval(context, data);
      
    case 'deny_device':
      return await handleDeviceDenial(context, data);
      
    default:
      return null;
  }
}

module.exports = {
  routeDeviceAction,
  handleDeviceApproval,
  handleDeviceDenial,
};
