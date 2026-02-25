/**
 * Share-related Adaptive Card builders
 * Cards for one-time share requests, search results, and status
 */

const { SHARE_DURATION_OPTIONS } = require('../constants');
const { buildSearchCardHeader, buildNoResultsSection } = require('../cardHelpers');
const { sanitizeHyperlinks } = require('../../utils/helpers');

/**
 * Build a search results card for one-time share
 */
function buildShareSearchResultsCard({
  approvalId,
  requesterName,
  requesterId,
  requesterEmail,
  requesterAadObjectId,
  justification,
  identifier,
  searchQuery,
  foundRecords,
  noResults = false,
  pamRecordsOnly = false,
  originalRecordTitle,
}) {
  const headerElements = buildSearchCardHeader('One-Time Share Request', requesterName, approvalId, justification);
  
  const card = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: [...headerElements],
    actions: [],
  };
  
  const baseData = { approvalId, identifier, requesterId, requesterEmail, requesterAadObjectId, requesterName, justification };
  
  if (noResults) {
    // Show specific message if all results were PAM records
    if (pamRecordsOnly) {
      card.body.push(
        { type: 'Container', style: 'attention', spacing: 'Medium', items: [
          { type: 'TextBlock', text: 'No Eligible Records Found', weight: 'Bolder', color: 'Attention' },
          { type: 'TextBlock', text: `Search: "${searchQuery}"`, size: 'Small', isSubtle: true, wrap: true },
        ]},
        { type: 'TextBlock', text: 'The search returned only PAM records (pamDirectory, pamDatabase, pamMachine, pamUser, pamRemoteBrowser).', wrap: true, spacing: 'Small' },
        { type: 'TextBlock', text: 'One-Time Shares are not available for PAM records. The requester should use `keeper-request-record` to request direct access instead.', wrap: true, spacing: 'Small', isSubtle: true },
        { type: 'TextBlock', text: 'Search Input', weight: 'Bolder', spacing: 'Medium' },
        { type: 'Input.Text', id: 'searchQuery', placeholder: 'Enter record name...', value: searchQuery || '' }
      );
    } else {
      card.body.push(...buildNoResultsSection(searchQuery, 'share'));
    }
    
    card.actions = [
      { type: 'Action.Execute', title: '🔍 Search', style: 'positive', verb: 'lookup_share', data: { action: 'lookup_share', ...baseData, recordTitle: originalRecordTitle } },
      { type: 'Action.Execute', title: 'Reset', verb: 'reset_share_card', data: { action: 'reset_share_card', ...baseData, recordTitle: originalRecordTitle } },
      { type: 'Action.Execute', title: 'Deny Request', style: 'destructive', verb: 'deny_share', data: { action: 'deny_share', approvalId, recordUid: null, recordTitle: originalRecordTitle, requesterId, requesterEmail, requesterName, justification } },
    ];
  } else if (foundRecords && foundRecords.length > 0) {
    const recordCount = foundRecords.length;
    
    if (recordCount === 1) {
      const record = foundRecords[0];
      card.body.push(
        { type: 'Container', style: 'good', spacing: 'Medium', items: [
          { type: 'TextBlock', text: `Record Found: ${record.title}`, wrap: true, weight: 'Bolder' },
          { type: 'TextBlock', text: `UID: ${record.uid}`, size: 'Small', isSubtle: true },
        ]},
        { type: 'TextBlock', text: 'Share Duration', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
        { type: 'Input.ChoiceSet', id: 'duration', value: '24h', choices: SHARE_DURATION_OPTIONS },
        { type: 'Input.Toggle', id: 'editable', title: 'Allow editing (recipient can modify the record)', value: 'false' }
      );
      
      card.actions = [
        { type: 'Action.Execute', title: 'Create Share', style: 'positive', verb: 'approve_share', data: { action: 'approve_share', approvalId, recordUid: record.uid, recordTitle: record.title, requesterId, requesterEmail, requesterName, justification } },
        { type: 'Action.Execute', title: 'Reset', verb: 'reset_share_card', data: { action: 'reset_share_card', ...baseData, recordTitle: originalRecordTitle } },
        { type: 'Action.Execute', title: 'Deny', style: 'destructive', verb: 'deny_share', data: { action: 'deny_share', approvalId, recordUid: record.uid, recordTitle: record.title, requesterId, requesterEmail, requesterName, justification } },
      ];
    } else {
      const recordChoices = foundRecords.map(r => ({ title: `${r.title} (${r.uid.substring(0, 8)}...)`, value: JSON.stringify({ uid: r.uid, title: r.title }) }));
      
      card.body.push(
        { type: 'Container', style: 'good', spacing: 'Medium', items: [
          { type: 'TextBlock', text: `${recordCount} Records Found`, wrap: true, weight: 'Bolder' },
          { type: 'TextBlock', text: 'Select the correct record from the list below:', size: 'Small', isSubtle: true },
        ]},
        { type: 'TextBlock', text: 'Select Record', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
        { type: 'Input.ChoiceSet', id: 'selectedRecord', value: recordChoices[0].value, choices: recordChoices, style: 'expanded' },
        { type: 'TextBlock', text: 'Share Duration', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
        { type: 'Input.ChoiceSet', id: 'duration', value: '24h', choices: SHARE_DURATION_OPTIONS },
        { type: 'Input.Toggle', id: 'editable', title: 'Allow editing (recipient can modify the record)', value: 'false' }
      );
      
      card.actions = [
        { type: 'Action.Execute', title: 'Create Share', style: 'positive', verb: 'approve_selected_share', data: { action: 'approve_selected_share', approvalId, requesterId, requesterEmail, requesterName, justification } },
        { type: 'Action.Execute', title: 'Reset', verb: 'reset_share_card', data: { action: 'reset_share_card', ...baseData, recordTitle: originalRecordTitle } },
        { type: 'Action.Execute', title: 'Deny Request', style: 'destructive', verb: 'deny_share', data: { action: 'deny_share', approvalId, recordUid: null, recordTitle: originalRecordTitle, requesterId, requesterEmail, requesterName, justification } },
      ];
    }
  }
  
  return card;
}

/**
 * Build an Adaptive Card for one-time share approval request
 */
function buildOneTimeShareApprovalCard({
  approvalId,
  requesterName,
  requesterId,
  requesterEmail,
  requesterAadObjectId,
  recordTitle,
  recordUid,
  justification,
  isUid = true,
  identifier,
}) {
  const requestedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
  
  // Sanitize identifier and justification to prevent URL injection
  const safeIdentifier = sanitizeHyperlinks(identifier || recordTitle);
  const safeJustification = sanitizeHyperlinks(justification) || 'No justification provided';
  
  const card = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      { type: 'TextBlock', text: 'One-Time Share Request', weight: 'Bolder', size: 'ExtraLarge' },
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              { type: 'TextBlock', text: 'Requester:', weight: 'Bolder', size: 'Medium' },
              { type: 'TextBlock', text: requesterName, color: 'Warning', size: 'Medium' },
              { type: 'TextBlock', text: 'Record:', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
              { type: 'TextBlock', text: safeIdentifier, color: 'Warning', size: 'Medium' },
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
              { type: 'TextBlock', text: safeJustification, wrap: true, size: 'Medium' },
            ],
          },
        ],
      },
    ],
    actions: [],
  };
  
  if (isUid && recordUid) {
    // Refresh property enables auto-refresh for all users when message is edited
    // Omitting userIds enables auto-refresh for ALL users in channels with <60 members
    card.refresh = {
      action: {
        type: 'Action.Execute',
        verb: 'refreshApprovalCard',
        data: {
          approvalId,
          type: 'share',
          requesterId,
          requesterEmail,
          requesterName,
          recordTitle,
          recordUid,
          justification,
          identifier,
          isUid,
        },
      },
    };
    // Add Record Details section when UID is resolved
    if (recordTitle && recordTitle !== identifier) {
      card.body.push({
        type: 'Container',
        separator: true,
        spacing: 'Medium',
        items: [
          { type: 'TextBlock', text: 'Record Details', weight: 'Bolder', size: 'Medium' },
          { type: 'FactSet', facts: [{ title: 'Title:', value: recordTitle }, { title: 'UID:', value: recordUid }] },
        ],
      });
    }
    
    // Duration and editable options
    card.body.push(
      { type: 'TextBlock', text: 'Share Duration', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
      { type: 'Input.ChoiceSet', id: 'duration', value: '24h', choices: SHARE_DURATION_OPTIONS },
      { type: 'Input.Toggle', id: 'editable', title: 'Allow editing (recipient can modify the record)', value: 'false' }
    );
    
    card.actions = [
      {
        type: 'Action.Execute',
        title: 'Approve',
        style: 'positive',
        verb: 'approve_share',
        data: { action: 'approve_share', approvalId, recordUid, recordTitle, requesterId, requesterEmail, requesterName, justification },
      },
      {
        type: 'Action.Execute',
        title: 'Deny',
        style: 'destructive',
        verb: 'deny_share',
        data: { action: 'deny_share', approvalId, recordUid, recordTitle, requesterId, requesterName },
      },
    ];
  } else {
    // Description-based - show inline search input
    card.body.push(
      { type: 'TextBlock', text: '**Action Required:** Search for the correct record to share', wrap: true, size: 'Medium', spacing: 'Large' },
      { type: 'Input.Text', id: 'searchQuery', placeholder: 'Enter record name or UID to search...', value: identifier || recordTitle || '' },
      { type: 'TextBlock', text: 'Enter a search term and click Search to find the record.', wrap: true, isSubtle: true, size: 'Small' }
    );
    
    card.actions = [
      {
        type: 'Action.Execute',
        title: '🔍 Search',
        style: 'positive',
        verb: 'lookup_share',
        data: { action: 'lookup_share', approvalId, identifier: identifier || recordTitle, recordTitle, requesterId, requesterEmail, requesterAadObjectId, requesterName, justification },
      },
      {
        type: 'Action.Execute',
        title: 'Deny Request',
        style: 'destructive',
        verb: 'deny_share',
        data: { action: 'deny_share', approvalId, recordUid: null, recordTitle: identifier || recordTitle, requesterId, requesterName },
      },
    ];
  }
  
  return card;
}

/**
 * Build a one-time share approval card with status (approved/denied)
 */
function buildOneTimeShareApprovalCardWithStatus({
  approvalId,
  requesterName,
  requesterEmail,
  recordTitle,
  recordUid,
  justification,
  status,
  approverName,
  duration,
  expiresAt,
  shareUrl,
  editable,
}) {
  const statusText = status === 'approved' ? 'APPROVED' : 'DENIED';
  const statusColor = status === 'approved' ? 'Good' : 'Attention';
  const time = new Date().toISOString().replace('T', ' ').substring(0, 19);
  
  const card = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: status === 'approved' ? 'good' : 'attention',
        items: [
          { type: 'TextBlock', text: 'One-Time Share Request', weight: 'Bolder', size: 'Large' },
          { type: 'TextBlock', text: statusText, weight: 'Bolder', color: statusColor },
        ],
      },
      {
        type: 'Container',
        items: [
          { type: 'FactSet', facts: [
            { title: 'Request ID', value: approvalId },
            { title: 'Requester', value: requesterName },
            { title: 'Record', value: recordTitle },
          ]},
        ],
      },
      {
        type: 'Container',
        items: [
          { type: 'TextBlock', text: 'Justification', weight: 'Bolder', size: 'Medium' },
          { type: 'TextBlock', text: justification || 'No justification provided', wrap: true },
        ],
      },
    ],
  };
  
  if (status === 'approved') {
    card.body.push({
      type: 'Container',
      separator: true,
      items: [
        { type: 'TextBlock', text: 'Share Details', weight: 'Bolder', size: 'Medium', color: 'Good' },
        { type: 'FactSet', facts: [
          { title: 'Reviewed by', value: approverName },
          { title: 'Duration', value: duration || '24h' },
          { title: 'Editable', value: editable ? 'Yes' : 'No' },
          { title: 'Expires', value: expiresAt || 'N/A' },
          { title: 'Processed', value: time },
        ]},
        { type: 'TextBlock', text: '💡 Share link has been sent to the requester via DM.', wrap: true, isSubtle: true },
      ],
    });
  } else {
    card.body.push({
      type: 'Container',
      separator: true,
      items: [
        { type: 'FactSet', facts: [
          { title: 'Denied by', value: approverName },
          { title: 'Time', value: time },
        ]},
      ],
    });
  }
  
  return card;
}

module.exports = {
  buildShareSearchResultsCard,
  buildOneTimeShareApprovalCard,
  buildOneTimeShareApprovalCardWithStatus,
};
