/**
 * Folder-related Adaptive Card builders
 * Cards for folder access requests, search results, and confirmations
 */

const { FOLDER_PERMISSIONS, DURATION_OPTIONS } = require('../constants');
const { 
  buildSearchCardHeader, 
  buildNoResultsSection, 
  buildFoundItemsHeader,
  formatFolderPermissionLabel 
} = require('../cardHelpers');

/**
 * Build an Adaptive Card for folder access approval request
 */
function buildFolderApprovalCard({
  approvalId,
  requesterName,
  requesterId,
  requesterEmail,
  requesterAadObjectId,
  folderName,
  folderUid,
  folderType = 'shared_folder',
  justification,
  isUid = true,
  identifier,
}) {
  const requestedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
  
  const card = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    refresh: {
      action: {
        type: 'Action.Execute',
        verb: 'refreshApprovalCard',
        data: {
          approvalId,
          type: 'folder',
          requesterId,
          requesterEmail,
          requesterName,
          folderName,
          folderUid,
          justification,
          identifier,
          isUid,
        },
      },
      userIds: [],
    },
    body: [
      { type: 'TextBlock', text: 'Folder Access Request', weight: 'Bolder', size: 'ExtraLarge' },
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              { type: 'TextBlock', text: 'Requester:', weight: 'Bolder', size: 'Medium' },
              { type: 'TextBlock', text: requesterName, color: 'Warning', size: 'Medium' },
              { type: 'TextBlock', text: 'Folder:', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
              { type: 'TextBlock', text: identifier || folderName, color: 'Warning', size: 'Medium' },
              { type: 'TextBlock', text: 'Requested:', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
              { type: 'TextBlock', text: requestedTime, size: 'Medium' },
            ],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              { type: 'TextBlock', text: 'Request ID:', weight: 'Bolder', size: 'Medium' },
              { type: 'TextBlock', text: approvalId, color: 'Warning', size: 'Medium' },
              { type: 'TextBlock', text: 'Justification:', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
              { type: 'TextBlock', text: justification || 'No justification provided', wrap: true, size: 'Medium' },
            ],
          },
        ],
      },
    ],
    actions: [],
  };
  
  if (isUid) {
    // Add Folder Details section when UID is resolved
    if (folderName && folderName !== identifier) {
      card.body.push({
        type: 'Container',
        separator: true,
        spacing: 'Medium',
        items: [
          { type: 'TextBlock', text: 'Folder Details', weight: 'Bolder', size: 'Medium' },
          { type: 'FactSet', facts: [{ title: 'Title:', value: folderName }, { title: 'Type:', value: folderType || 'Shared Folder' }] },
        ],
      });
    }
    
    // Permission/duration selectors
    card.body.push(
      { type: 'TextBlock', text: 'Permission Level', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
      { type: 'Input.ChoiceSet', id: 'permission', value: 'no_permissions', choices: FOLDER_PERMISSIONS },
      { type: 'TextBlock', text: 'Duration', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
      { type: 'Input.ChoiceSet', id: 'duration', value: '24h', choices: DURATION_OPTIONS }
    );
    
    card.actions = [
      {
        type: 'Action.Execute',
        title: 'Approve',
        style: 'positive',
        verb: 'approve_folder',
        data: { action: 'approve_folder', approvalId, folderUid, folderName, requesterId, requesterEmail, requesterName },
      },
      {
        type: 'Action.Execute',
        title: 'Deny',
        style: 'destructive',
        verb: 'deny_folder',
        data: { action: 'deny_folder', approvalId, folderUid, folderName, requesterId, requesterName },
      },
    ];
  } else {
    // Description - show inline search input
    card.body.push(
      { type: 'TextBlock', text: '**Action Required:** Search for the correct folder', wrap: true, size: 'Medium', spacing: 'Large' },
      { type: 'Input.Text', id: 'searchQuery', placeholder: 'Enter folder name or UID to search...', value: identifier || folderName || '' },
      { type: 'TextBlock', text: 'Enter a search term and click Look Up to find the folder.', wrap: true, isSubtle: true, size: 'Small' }
    );
    
    card.actions = [
      {
        type: 'Action.Execute',
        title: '🔍 Look Up',
        style: 'positive',
        verb: 'lookup_folder',
        data: { action: 'lookup_folder', approvalId, identifier: identifier || folderName, folderName, requesterId, requesterEmail, requesterAadObjectId, requesterName, justification },
      },
      {
        type: 'Action.Execute',
        title: 'Deny Request',
        style: 'destructive',
        verb: 'deny_folder',
        data: { action: 'deny_folder', approvalId, folderUid: null, folderName, requesterId, requesterEmail, requesterName, justification },
      },
    ];
  }
  
  return card;
}

/**
 * Build a folder search results card (inline search flow)
 */
function buildFolderSearchResultsCard({
  approvalId,
  requesterName,
  requesterId,
  requesterEmail,
  requesterAadObjectId,
  justification,
  identifier,
  searchQuery,
  foundFolders,
  noResults = false,
  originalFolderName,
}) {
  const headerElements = buildSearchCardHeader('Folder Access Request', requesterName, approvalId, justification);
  
  const card = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: [...headerElements],
    actions: [],
  };
  
  const baseData = { approvalId, identifier, requesterId, requesterEmail, requesterAadObjectId, requesterName, justification };
  
  if (noResults) {
    card.body.push(...buildNoResultsSection(searchQuery, 'folder'));
    
    card.actions = [
      { type: 'Action.Execute', title: '🔍 Look Up', style: 'positive', verb: 'lookup_folder', data: { action: 'lookup_folder', ...baseData, folderName: originalFolderName } },
      { type: 'Action.Execute', title: 'Reset', verb: 'reset_folder_card', data: { action: 'reset_folder_card', ...baseData, folderName: originalFolderName } },
      { type: 'Action.Execute', title: 'Deny Request', style: 'destructive', verb: 'deny_folder', data: { action: 'deny_folder', approvalId, folderUid: null, folderName: originalFolderName, requesterId, requesterEmail, requesterName, justification } },
    ];
  } else if (foundFolders && foundFolders.length > 0) {
    const folderCount = foundFolders.length;
    const singleFolder = folderCount === 1 ? { uid: foundFolders[0].uid, title: foundFolders[0].name } : null;
    card.body.push(...buildFoundItemsHeader(folderCount, 'folder', singleFolder));
    
    if (folderCount === 1) {
      const folder = foundFolders[0];
      card.body.push(
        { type: 'TextBlock', text: 'Permission Level', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
        { type: 'Input.ChoiceSet', id: 'permission', value: 'no_permissions', choices: FOLDER_PERMISSIONS },
        { type: 'TextBlock', text: 'Duration', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
        { type: 'Input.ChoiceSet', id: 'duration', value: '1h', choices: DURATION_OPTIONS }
      );
      
      card.actions = [
        { type: 'Action.Execute', title: 'Approve', style: 'positive', verb: 'approve_folder', data: { action: 'approve_folder', approvalId, folderUid: folder.uid, folderName: folder.name, requesterId, requesterEmail, requesterName, justification } },
        { type: 'Action.Execute', title: 'Reset', verb: 'reset_folder_card', data: { action: 'reset_folder_card', ...baseData, folderName: originalFolderName } },
        { type: 'Action.Execute', title: 'Deny', style: 'destructive', verb: 'deny_folder', data: { action: 'deny_folder', approvalId, folderUid: folder.uid, folderName: folder.name, requesterId, requesterEmail, requesterName, justification } },
      ];
    } else {
      const folderChoices = foundFolders.map(f => ({ title: `${f.name} (${f.uid.substring(0, 8)}...)`, value: JSON.stringify({ uid: f.uid, name: f.name }) }));
      
      card.body.push(
        { type: 'TextBlock', text: 'Select Folder', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
        { type: 'Input.ChoiceSet', id: 'selectedFolder', value: folderChoices[0].value, choices: folderChoices, style: 'expanded' },
        { type: 'TextBlock', text: 'Permission Level', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
        { type: 'Input.ChoiceSet', id: 'permission', value: 'no_permissions', choices: FOLDER_PERMISSIONS },
        { type: 'TextBlock', text: 'Duration', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
        { type: 'Input.ChoiceSet', id: 'duration', value: '1h', choices: DURATION_OPTIONS }
      );
      
      card.actions = [
        { type: 'Action.Execute', title: 'Approve Selected', style: 'positive', verb: 'approve_selected_folder', data: { action: 'approve_selected_folder', approvalId, requesterId, requesterEmail, requesterName, justification } },
        { type: 'Action.Execute', title: 'Reset', verb: 'reset_folder_card', data: { action: 'reset_folder_card', ...baseData, folderName: originalFolderName } },
        { type: 'Action.Execute', title: 'Deny Request', style: 'destructive', verb: 'deny_folder', data: { action: 'deny_folder', approvalId, folderUid: null, folderName: originalFolderName, requesterId, requesterEmail, requesterName, justification } },
      ];
    }
  }
  
  return card;
}

/**
 * Build a folder approval card with status (for updating existing card)
 */
function buildFolderApprovalCardWithStatus({
  approvalId,
  requesterName,
  requesterEmail,
  folderName,
  justification,
  status,
  approverName,
  permission,
  duration,
  expiresAt,
  processedTime,
}) {
  const statusText = status === 'approved' ? 'APPROVED' : 'DENIED';
  const time = processedTime || new Date().toISOString().replace('T', ' ').substring(0, 19);
  
  const card = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.2',
    body: [
      { type: 'TextBlock', text: 'Folder Access Request', weight: 'Bolder', size: 'ExtraLarge' },
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              { type: 'TextBlock', text: 'Requester:', weight: 'Bolder', size: 'Medium' },
              { type: 'TextBlock', text: requesterName || 'Unknown', color: 'Warning', size: 'Medium' },
              { type: 'TextBlock', text: 'Folder:', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
              { type: 'TextBlock', text: folderName || 'Unknown', color: 'Warning', size: 'Medium' },
            ],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              { type: 'TextBlock', text: 'Request ID:', weight: 'Bolder', size: 'Medium' },
              { type: 'TextBlock', text: approvalId || 'N/A', color: 'Warning', size: 'Medium' },
              { type: 'TextBlock', text: 'Justification:', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
              { type: 'TextBlock', text: justification || 'No justification provided', wrap: true, size: 'Medium' },
            ],
          },
        ],
      },
      {
        type: 'Container',
        style: status === 'approved' ? 'good' : 'attention',
        items: [
          { type: 'TextBlock', text: statusText, weight: 'Bolder', size: 'Large', horizontalAlignment: 'Center' },
          { type: 'TextBlock', text: `By: ${approverName || 'Unknown'} at ${time}`, size: 'Small', horizontalAlignment: 'Center', isSubtle: true },
        ],
      },
    ],
    actions: [],
  };
  
  if (status === 'approved') {
    const detailsItems = [];
    
    if (folderName) {
      detailsItems.push({
        type: 'ColumnSet',
        columns: [
          { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: 'Access granted for:', weight: 'Bolder', size: 'Small' }] },
          { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: folderName, size: 'Small', color: 'Good' }] },
        ],
      });
    }
    
    if (permission) {
      detailsItems.push({
        type: 'ColumnSet',
        columns: [
          { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: 'Permission:', weight: 'Bolder', size: 'Small' }] },
          { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: formatFolderPermissionLabel(permission), size: 'Small' }] },
        ],
      });
    }
    
    if (duration) {
      detailsItems.push({
        type: 'ColumnSet',
        columns: [
          { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: 'Duration:', weight: 'Bolder', size: 'Small' }] },
          { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: duration, size: 'Small' }] },
        ],
      });
    }
    
    if (requesterEmail) {
      detailsItems.push({
        type: 'ColumnSet',
        columns: [
          { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: 'Granted To:', weight: 'Bolder', size: 'Small' }] },
          { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: requesterEmail, size: 'Small' }] },
        ],
      });
    }
    
    if (expiresAt) {
      detailsItems.push({
        type: 'ColumnSet',
        columns: [
          { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: 'Expires:', weight: 'Bolder', size: 'Small' }] },
          { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: expiresAt, size: 'Small' }] },
        ],
      });
    }
    
    if (detailsItems.length > 0) {
      card.body.splice(2, 0, { type: 'Container', spacing: 'Medium', items: detailsItems });
    }
  }
  
  return card;
}

/**
 * Build a folder confirmation card (after search/selection in task module)
 */
function buildFolderConfirmationCard({
  approvalId,
  requesterName,
  requesterId,
  requesterEmail,
  folderName,
  folderUid,
  justification,
  permission,
  duration,
}) {
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: [
      { type: 'TextBlock', text: 'Folder Selected - Ready to Approve', weight: 'Bolder', size: 'Large', color: 'Good' },
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              { type: 'TextBlock', text: 'Requester:', weight: 'Bolder', size: 'Medium' },
              { type: 'TextBlock', text: requesterName || 'Unknown', color: 'Warning', size: 'Medium' },
              { type: 'TextBlock', text: 'Selected Folder:', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
              { type: 'TextBlock', text: folderName || folderUid, color: 'Good', size: 'Medium', weight: 'Bolder' },
            ],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              { type: 'TextBlock', text: 'Request ID:', weight: 'Bolder', size: 'Medium' },
              { type: 'TextBlock', text: approvalId || 'N/A', color: 'Warning', size: 'Medium' },
              { type: 'TextBlock', text: 'Justification:', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
              { type: 'TextBlock', text: justification || 'No justification provided', wrap: true, size: 'Medium' },
            ],
          },
        ],
      },
      {
        type: 'Container',
        separator: true,
        spacing: 'Medium',
        items: [
          { type: 'TextBlock', text: 'Access Configuration', weight: 'Bolder', size: 'Medium' },
          { type: 'FactSet', facts: [
            { title: 'Permission:', value: formatFolderPermissionLabel(permission) },
            { title: 'Duration:', value: duration === 'permanent' ? 'Permanent' : duration },
            { title: 'For User:', value: requesterEmail },
          ]},
        ],
      },
      { type: 'TextBlock', text: 'Click **Approve** to grant access or **Deny** to reject this request.', wrap: true, isSubtle: true, size: 'Small', spacing: 'Medium' },
    ],
    actions: [
      {
        type: 'Action.Execute',
        title: 'Approve',
        style: 'positive',
        verb: 'approve_folder',
        data: { action: 'approve_folder', approvalId, folderUid, folderName, requesterId, requesterEmail, requesterName, justification, permission, duration },
      },
      {
        type: 'Action.Execute',
        title: 'Deny',
        style: 'destructive',
        verb: 'deny_folder',
        data: { action: 'deny_folder', approvalId, folderUid, folderName, requesterId, requesterEmail, requesterName, justification },
      },
    ],
  };
}

module.exports = {
  buildFolderApprovalCard,
  buildFolderSearchResultsCard,
  buildFolderApprovalCardWithStatus,
  buildFolderConfirmationCard,
};
