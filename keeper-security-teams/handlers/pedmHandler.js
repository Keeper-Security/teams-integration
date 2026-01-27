/**
 * PEDM (Privileged Elevation & Delegation Management) Handler
 * 
 * Handles Adaptive Card action submissions for PEDM elevation requests:
 * - Approve PEDM request
 * - Deny PEDM request
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
 * Handle approval of a PEDM elevation request
 */
async function handlePedmApproval(context, data) {
  const approver = getApproverInfo(context.activity);
  const approvalUid = data.approvalUid;
  const username = data.username || 'Unknown User';
  const command = data.command || 'Unknown Command';
  
  if (!approvalUid) {
    await context.send('❌ Error: Missing approval UID');
    return;
  }
  
  await context.send('🔄 Approving PEDM request...');
  
  const result = await keeperClient.approvePedmRequest(approvalUid);
  
  if (result.success) {
    const approvedCard = cards.buildPedmApprovedCard(
      approver.name,
      username,
      command
    );
    
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: approvedCard,
      }],
    });
    
    await context.send('✅ PEDM request approved for **' + username + '**');
  } else {
    await context.send('❌ Failed to approve PEDM request: ' + result.error);
  }
}

/**
 * Handle denial of a PEDM elevation request
 */
async function handlePedmDenial(context, data) {
  const approver = getApproverInfo(context.activity);
  const approvalUid = data.approvalUid;
  const username = data.username || 'Unknown User';
  const command = data.command || 'Unknown Command';
  
  if (!approvalUid) {
    await context.send('❌ Error: Missing approval UID');
    return;
  }
  
  await context.send('🔄 Denying PEDM request...');
  
  const result = await keeperClient.denyPedmRequest(approvalUid);
  
  if (result.success) {
    const deniedCard = cards.buildPedmDeniedCard(
      approver.name,
      username,
      command
    );
    
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: deniedCard,
      }],
    });
    
    await context.send('❌ PEDM request denied for **' + username + '**');
  } else {
    await context.send('❌ Failed to deny PEDM request: ' + result.error);
  }
}

/**
 * Route PEDM card action to appropriate handler
 */
async function routePedmAction(context, data) {
  const action = data.action;
  
  switch (action) {
    case 'approve_pedm':
      await handlePedmApproval(context, data);
      return true;
      
    case 'deny_pedm':
      await handlePedmDenial(context, data);
      return true;
      
    default:
      return false;
  }
}

module.exports = {
  routePedmAction,
  handlePedmApproval,
  handlePedmDenial,
};
