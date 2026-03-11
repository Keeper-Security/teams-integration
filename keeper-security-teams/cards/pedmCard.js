/**
 * Adaptive Card builders for PEDM (Privileged Elevation & Delegation Management)
 * 
 * These cards are displayed when PEDM elevation requests are detected,
 * allowing admins to approve or deny privilege escalation requests.
 */

/**
 * Build an Adaptive Card for PEDM elevation request
 * Includes all fields including agentUid
 */
function buildPedmApprovalCard({
  approvalUid,
  approvalType,
  agentUid,
  username,
  command,
  fileName,
  filePath,
  description,
  justification,
  expireIn,
  created,
}) {
  const title = 'Endpoint Privilege Manage Request';
  
  const facts = [
    { title: 'User', value: username || 'Unknown' },
    { title: 'Type', value: approvalType || 'Unknown' },
  ];
  
  // Add Agent UID if available
  if (agentUid) {
    facts.push({ title: 'Agent UID', value: agentUid });
  }
  
  if (fileName) {
    facts.push({ title: 'File', value: fileName });
  }
  
  if (filePath) {
    facts.push({ title: 'Path', value: filePath });
  }
  
  if (expireIn) {
    facts.push({ title: 'Expires In', value: expireIn + ' minutes' });
  }
  
  if (created) {
    facts.push({ title: 'Requested', value: formatDate(created) });
  }
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    // Refresh property enables auto-refresh for all users when message is edited
    // Omitting userIds enables auto-refresh for ALL users in channels with <60 members
    refresh: {
      action: {
        type: 'Action.Execute',
        verb: 'refreshPedmCard',
        data: {
          action: 'refreshPedmCard',
          approvalUid: approvalUid,
          agentUid: agentUid,
          username: username,
          command: command || fileName,
        },
      },
    },
    body: [
      {
        type: 'Container',
        style: 'warning',
        items: [
          { type: 'TextBlock', text: title, weight: 'Bolder', size: 'Large', color: 'Warning' },
          { type: 'TextBlock', text: 'UID: ' + approvalUid, size: 'Small', isSubtle: true },
        ],
      },
      {
        type: 'Container',
        items: [
          { type: 'FactSet', facts: facts },
        ],
      },
      ...(command ? [{
        type: 'Container',
        items: [
          { type: 'TextBlock', text: 'Command', weight: 'Bolder' },
          { type: 'TextBlock', text: '`' + command + '`', wrap: true, fontType: 'Monospace' },
        ],
      }] : []),
      ...(description ? [{
        type: 'Container',
        items: [
          { type: 'TextBlock', text: 'Description', weight: 'Bolder' },
          { type: 'TextBlock', text: description, wrap: true },
        ],
      }] : []),
      ...(justification ? [{
        type: 'Container',
        items: [
          { type: 'TextBlock', text: 'Justification', weight: 'Bolder' },
          { type: 'TextBlock', text: justification, wrap: true },
        ],
      }] : []),
    ],
    actions: [
      {
        type: 'Action.Execute',
        title: 'Approve',
        style: 'positive',
        verb: 'approve_pedm',
        data: {
          action: 'approve_pedm',
          approvalUid: approvalUid,
          agentUid: agentUid,
          username: username,
          command: command || fileName,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny',
        style: 'destructive',
        verb: 'deny_pedm',
        data: {
          action: 'deny_pedm',
          approvalUid: approvalUid,
          agentUid: agentUid,
          username: username,
          command: command || fileName,
        },
      },
    ],
  };
}

/**
 * Build an Adaptive Card showing PEDM request was approved
 * Updated to include approvalUid and agentUid
 */
function buildPedmApprovedCard(approverName, username, command, approvalUid, agentUid) {
  const facts = [
    { title: 'User', value: username },
    { title: 'Approved by', value: approverName },
  ];
  
  if (approvalUid) {
    facts.push({ title: 'Approval UID', value: approvalUid });
  }
  
  if (agentUid) {
    facts.push({ title: 'Agent UID', value: agentUid });
  }
  
  if (command) {
    facts.push({ title: 'Command/File', value: command });
  }
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: 'good',
        items: [
          { type: 'TextBlock', text: 'EPM Request Approved', weight: 'Bolder', size: 'Medium', color: 'Good' },
        ],
      },
      {
        type: 'Container',
        items: [
          { type: 'FactSet', facts: facts },
        ],
      },
    ],
  };
}

/**
 * Build an Adaptive Card showing PEDM request was denied
 * Updated to include approvalUid and agentUid
 */
function buildPedmDeniedCard(approverName, username, command, approvalUid, agentUid) {
  const facts = [
    { title: 'User', value: username },
    { title: 'Denied by', value: approverName },
  ];
  
  if (approvalUid) {
    facts.push({ title: 'Approval UID', value: approvalUid });
  }
  
  if (agentUid) {
    facts.push({ title: 'Agent UID', value: agentUid });
  }
  
  if (command) {
    facts.push({ title: 'Command/File', value: command });
  }
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: 'attention',
        items: [
          { type: 'TextBlock', text: 'EPM Request Denied', weight: 'Bolder', size: 'Medium', color: 'Attention' },
        ],
      },
      {
        type: 'Container',
        items: [
          { type: 'FactSet', facts: facts },
        ],
      },
    ],
  };
}

function formatDate(dateStr) {
  try {
    const date = new Date(dateStr);
    return date.toLocaleString('en-US', {
      month: '2-digit',
      day: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true,
    });
  } catch (e) {
    return dateStr;
  }
}

module.exports = {
  buildPedmApprovalCard,
  buildPedmApprovedCard,
  buildPedmDeniedCard,
};
