/**
 * Device Approval Handler
 * 
 * Handles Adaptive Card action submissions for Cloud SSO device approvals:
 * - Approve device
 * - Deny device
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
 * Handle approval of a device request
 */
async function handleDeviceApproval(context, data) {
  const approver = getApproverInfo(context.activity);
  const deviceId = data.deviceId;
  const deviceName = data.deviceName || 'Unknown Device';
  const username = data.username || 'Unknown User';
  
  if (!deviceId) {
    await context.send('❌ Error: Missing device ID');
    return;
  }
  
  await context.send('🔄 Approving device...');
  
  const result = await keeperClient.approveDevice(deviceId);
  
  if (result.success) {
    const approvedCard = cards.buildDeviceApprovedCard(
      approver.name,
      deviceName,
      username
    );
    
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: approvedCard,
      }],
    });
    
    await context.send('✅ Device approved for **' + username + '**');
  } else if (result.alreadyHandled) {
    await context.send('⚠️ This device request was already processed.');
  } else {
    await context.send('❌ Failed to approve device: ' + result.error);
  }
}

/**
 * Handle denial of a device request
 */
async function handleDeviceDenial(context, data) {
  const approver = getApproverInfo(context.activity);
  const deviceId = data.deviceId;
  const deviceName = data.deviceName || 'Unknown Device';
  const username = data.username || 'Unknown User';
  
  if (!deviceId) {
    await context.send('❌ Error: Missing device ID');
    return;
  }
  
  await context.send('🔄 Denying device...');
  
  const result = await keeperClient.denyDevice(deviceId);
  
  if (result.success) {
    const deniedCard = cards.buildDeviceDeniedCard(
      approver.name,
      deviceName,
      username
    );
    
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: deniedCard,
      }],
    });
    
    await context.send('❌ Device denied for **' + username + '**');
  } else if (result.alreadyHandled) {
    await context.send('⚠️ This device request was already processed.');
  } else {
    await context.send('❌ Failed to deny device: ' + result.error);
  }
}

/**
 * Route device card action to appropriate handler
 */
async function routeDeviceAction(context, data) {
  const action = data.action;
  
  switch (action) {
    case 'approve_device':
      await handleDeviceApproval(context, data);
      return true;
      
    case 'deny_device':
      await handleDeviceDenial(context, data);
      return true;
      
    default:
      return false;
  }
}

module.exports = {
  routeDeviceAction,
  handleDeviceApproval,
  handleDeviceDenial,
};
