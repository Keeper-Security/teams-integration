/**
 * Adaptive Card builders for PEDM (Privileged Elevation & Delegation Management)
 * 
 * These cards are displayed when PEDM elevation requests are detected,
 * allowing admins to approve or deny privilege escalation requests.
 */

/**
 * Build an Adaptive Card for PEDM elevation request
 */
function buildPedmApprovalCard({
  approvalUid,
  approvalType,
  username,
  command,
  fileName,
  filePath,
  description,
  justification,
  expireIn,
  created,
}) {
  const isCommandLine = approvalType === 'CommandLine';
  const icon = isCommandLine ? '⌨️' : '🔓';
  const title = isCommandLine ? 'Command Execution Request' : 'Privilege Elevation Request';
  
  const facts = [
    { title: 'User', value: username || 'Unknown' },
    { title: 'Type', value: approvalType || 'Unknown' },
  ];
  
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
    body: [
      {
        type: 'Container',
        style: 'warning',
        items: [
          {
            type: 'ColumnSet',
            columns: [
              {
                type: 'Column',
                width: 'auto',
                items: [{ type: 'TextBlock', text: icon, size: 'ExtraLarge' }],
              },
              {
                type: 'Column',
                width: 'stretch',
                items: [
                  { type: 'TextBlock', text: title, weight: 'Bolder', size: 'Large', color: 'Warning' },
                  { type: 'TextBlock', text: 'UID: ' + approvalUid, size: 'Small', isSubtle: true },
                ],
              },
            ],
          },
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
        type: 'Action.Submit',
        title: '✅ Approve',
        style: 'positive',
        data: {
          action: 'approve_pedm',
          approvalUid: approvalUid,
          username: username,
          command: command || fileName,
        },
      },
      {
        type: 'Action.Submit',
        title: '❌ Deny',
        style: 'destructive',
        data: {
          action: 'deny_pedm',
          approvalUid: approvalUid,
          username: username,
          command: command || fileName,
        },
      },
    ],
  };
}

/**
 * Build an Adaptive Card showing PEDM request was approved
 */
function buildPedmApprovedCard(approverName, username, command) {
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: 'good',
        items: [
          {
            type: 'ColumnSet',
            columns: [
              {
                type: 'Column',
                width: 'auto',
                items: [{ type: 'TextBlock', text: '✅', size: 'Large' }],
              },
              {
                type: 'Column',
                width: 'stretch',
                items: [
                  { type: 'TextBlock', text: 'PEDM Request Approved', weight: 'Bolder', size: 'Medium', color: 'Good' },
                ],
              },
            ],
          },
        ],
      },
      {
        type: 'Container',
        items: [
          {
            type: 'FactSet',
            facts: [
              { title: 'User', value: username },
              { title: 'Approved by', value: approverName },
              ...(command ? [{ title: 'Command/File', value: command }] : []),
            ],
          },
        ],
      },
    ],
  };
}

/**
 * Build an Adaptive Card showing PEDM request was denied
 */
function buildPedmDeniedCard(approverName, username, command) {
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: 'attention',
        items: [
          {
            type: 'ColumnSet',
            columns: [
              {
                type: 'Column',
                width: 'auto',
                items: [{ type: 'TextBlock', text: '❌', size: 'Large' }],
              },
              {
                type: 'Column',
                width: 'stretch',
                items: [
                  { type: 'TextBlock', text: 'PEDM Request Denied', weight: 'Bolder', size: 'Medium', color: 'Attention' },
                ],
              },
            ],
          },
        ],
      },
      {
        type: 'Container',
        items: [
          {
            type: 'FactSet',
            facts: [
              { title: 'User', value: username },
              { title: 'Denied by', value: approverName },
              ...(command ? [{ title: 'Command/File', value: command }] : []),
            ],
          },
        ],
      },
    ],
  };
}

function formatDate(dateStr) {
  try {
    const date = new Date(dateStr);
    return date.toLocaleString();
  } catch (e) {
    return dateStr;
  }
}

module.exports = {
  buildPedmApprovalCard,
  buildPedmApprovedCard,
  buildPedmDeniedCard,
};
