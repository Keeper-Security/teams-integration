/**
 * Device Approval Handler
 * 
 * Handles Adaptive Card action submissions for Cloud SSO device approvals:
 * - Approve device
 * - Deny device
 * 
 * Features:
 * - In-place card updates (not new messages)
 * - Already processed request handling
 * - Channel card updates for all users via message edit
 */

const keeperClient = require('../services/keeperClient');
const cards = require('../cards');
const { createLogger, getChannelService } = require('../services');
const { getCurrentTimestamp } = require('./approval/helpers');

const log = createLogger('DeviceHandler');

/**
 * Helper function to update a device approval card in the channel
 * Uses context.updateActivity() first (most reliable in invoke handlers),
 * then falls back to channelService.updateApprovalCard()
 * @param {string} deviceId - The device ID
 * @param {Object} updatedCard - The updated Adaptive Card content
 * @param {Object} context - The Teams context
 */
async function tryUpdateDeviceCard(deviceId, updatedCard, context) {
  // replyToId is the activity ID of the card being interacted with
  const activityId = context?.activity?.replyToId;
  
  if (!activityId) {
    log.debug(`No replyToId found for device ${deviceId}`);
    return false;
  }

  log.debug(`Updating device card ${deviceId}`, { activityId });

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
      log.info('Updated device card via context.updateActivity()', { deviceId, activityId });
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
    log.info('Updated device card via channelService', { deviceId, activityId });
    return true;
  }

  log.debug('Failed to update device card in channel');
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
 * Build an "already processed" card when the device request was handled elsewhere
 */
function buildAlreadyProcessedCard(username, deviceId, checkedBy) {
  const timestamp = getCurrentTimestamp();
  
  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
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
        text: 'This device approval request for **' + username + '** has already been approved or denied from another platform (Keeper Vault or another integration).',
        wrap: true,
        spacing: 'Medium',
      },
      {
        type: 'FactSet',
        facts: [
          { title: 'Device ID', value: deviceId },
          { title: 'Status', value: 'Already Processed (approved/denied elsewhere)' },
          { title: 'Checked by', value: checkedBy || 'Admin' },
          { title: 'Updated', value: timestamp },
        ],
        spacing: 'Medium',
      },
    ],
  };
}

/**
 * Handle approval of a device request
 * Returns Adaptive Card for in-place update and updates channel card for all users
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
      version: '1.5',
      body: [{
        type: 'TextBlock',
        text: 'Error: Missing device ID',
        color: 'Attention',
      }],
    };
  }
  
  log.info('Approving device', { deviceId, deviceName, username, approver: approver.name });
  
  const result = await keeperClient.approveDevice(deviceId);
  
  log.info('Device approval result', { deviceId, success: result.success, already_processed: result.already_processed, error: result.error });
  
  let updatedCard;
  
  if (result.success) {
    log.info('Device approved successfully', { deviceId, approver: approver.name });
    updatedCard = cards.buildDeviceApprovedCard(
      approver.name,
      deviceName,
      username,
      deviceId
    );
  } else if (result.already_processed) {
    log.info('Device already processed elsewhere', { deviceId });
    updatedCard = buildAlreadyProcessedCard(username, deviceId, approver.name);
  } else {
    // Return error card
    updatedCard = {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.5',
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
  
  // Update the channel card for all users (non-blocking to avoid Teams timeout)
  tryUpdateDeviceCard(deviceId, updatedCard, context)
    .then(() => log.debug('Updated device card via channelService'))
    .catch(err => log.debug('Could not update channel card', err.message));
  
  return updatedCard;
}

/**
 * Handle denial of a device request
 * Returns Adaptive Card for in-place update and updates channel card for all users
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
      version: '1.5',
      body: [{
        type: 'TextBlock',
        text: 'Error: Missing device ID',
        color: 'Attention',
      }],
    };
  }
  
  log.info('Denying device', { deviceId, deviceName, username, approver: approver.name });
  
  const result = await keeperClient.denyDevice(deviceId);
  
  log.info('Device denial result', { deviceId, success: result.success, already_processed: result.already_processed, error: result.error });
  
  let updatedCard;
  
  if (result.success) {
    log.info('Device denied successfully', { deviceId, approver: approver.name });
    updatedCard = cards.buildDeviceDeniedCard(
      approver.name,
      deviceName,
      username,
      deviceId
    );
  } else if (result.already_processed) {
    log.info('Device already processed elsewhere', { deviceId });
    updatedCard = buildAlreadyProcessedCard(username, deviceId, approver.name);
  } else {
    // Return error card
    updatedCard = {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.5',
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
  
  // Update the channel card for all users (non-blocking to avoid Teams timeout)
  tryUpdateDeviceCard(deviceId, updatedCard, context)
    .then(() => log.debug('Updated device card via channelService'))
    .catch(err => log.debug('Could not update channel card', err.message));
  
  return updatedCard;
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
