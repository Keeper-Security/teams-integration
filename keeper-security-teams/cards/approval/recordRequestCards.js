/**
 * Record-related Adaptive Card builders
 * Cards for record access requests, search results, and confirmations
 */

const { RECORD_PERMISSIONS, DURATION_OPTIONS, SELF_DESTRUCT_DURATION_OPTIONS } = require('../constants');
const { 
  buildSearchCardHeader, 
  buildNoResultsSection, 
  formatPermissionLabel 
} = require('../cardHelpers');

/**
 * Build an Adaptive Card for record access approval request
 */
function buildRecordApprovalCard({
  approvalId,
  requesterName,
  requesterId,
  requesterEmail,
  requesterAadObjectId,
  recordTitle,
  recordUid,
  recordType = 'login',
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
          type: 'record',
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
      userIds: [],
    },
    body: [
      { type: 'TextBlock', text: 'Record Access Request', weight: 'Bolder', size: 'ExtraLarge' },
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
              { type: 'TextBlock', text: identifier || recordTitle, color: 'Warning', size: 'Medium' },
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
    // Add Record Details section when UID is resolved
    if (recordTitle && recordTitle !== identifier) {
      card.body.push({
        type: 'Container',
        separator: true,
        spacing: 'Medium',
        items: [
          { type: 'TextBlock', text: 'Record Details', weight: 'Bolder', size: 'Medium' },
          { type: 'FactSet', facts: [{ title: 'Title:', value: recordTitle }, { title: 'Type:', value: recordType || 'Record' }] },
        ],
      });
    }
    
    // Permission/duration selectors
    card.body.push(
      { type: 'TextBlock', text: 'Permission Level', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
      { type: 'Input.ChoiceSet', id: 'permission', value: 'view_only', choices: RECORD_PERMISSIONS },
      { type: 'TextBlock', text: 'Duration', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
      { type: 'Input.ChoiceSet', id: 'duration', value: '24h', choices: DURATION_OPTIONS }
    );
    
    card.actions = [
      {
        type: 'Action.Execute',
        title: 'Approve',
        style: 'positive',
        verb: 'approve_record',
        data: { action: 'approve_record', approvalId, recordUid, recordTitle, requesterId, requesterEmail, requesterName },
      },
      {
        type: 'Action.Execute',
        title: 'Deny',
        style: 'destructive',
        verb: 'deny_record',
        data: { action: 'deny_record', approvalId, recordUid, recordTitle, requesterId, requesterName },
      },
    ];
  } else {
    // Description - show inline search input
    card.body.push(
      { type: 'TextBlock', text: '**Action Required:** Search for the correct record', wrap: true, size: 'Medium', spacing: 'Large' },
      { type: 'Input.Text', id: 'searchQuery', placeholder: 'Enter record name or UID to search...', value: identifier || recordTitle || '' },
      { type: 'TextBlock', text: 'Enter a search term and click Look Up to find the record.', wrap: true, isSubtle: true, size: 'Small' }
    );
    
    card.actions = [
      {
        type: 'Action.Execute',
        title: '🔍 Search',
        style: 'positive',
        verb: 'lookup_record',
        data: { action: 'lookup_record', approvalId, identifier: identifier || recordTitle, recordTitle, requesterId, requesterEmail, requesterAadObjectId, requesterName, justification },
      },
      {
        type: 'Action.Execute',
        title: 'Create New Record',
        verb: 'show_create_form',
        data: { action: 'show_create_form', approvalId, identifier: identifier || recordTitle, recordTitle, requesterId, requesterEmail, requesterAadObjectId, requesterName, justification, searchQuery: identifier || recordTitle },
      },
      {
        type: 'Action.Execute',
        title: 'Deny Request',
        style: 'destructive',
        verb: 'deny_record',
        data: { action: 'deny_record', approvalId, recordUid: null, recordTitle, requesterId, requesterEmail, requesterName, justification },
      },
    ];
  }
  
  return card;
}

/**
 * Build a record search results card (inline search flow)
 */
function buildRecordSearchResultsCard({
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
  originalRecordTitle,
}) {
  const headerElements = buildSearchCardHeader('Record Access Request', requesterName, approvalId, justification);
  
  const card = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: [...headerElements],
    actions: [],
  };
  
  const baseData = { approvalId, identifier, requesterId, requesterEmail, requesterAadObjectId, requesterName, justification };
  
  if (noResults) {
    card.body.push(...buildNoResultsSection(searchQuery, 'record'));
    
    // Add hint about creating new record
    card.body.push({
      type: 'TextBlock',
      text: 'Or create a new record to share:',
      wrap: true,
      spacing: 'Medium',
      isSubtle: true,
    });
    
    card.actions = [
      { type: 'Action.Execute', title: '🔍 Search', style: 'positive', verb: 'lookup_record', data: { action: 'lookup_record', ...baseData, recordTitle: originalRecordTitle } },
      { 
        type: 'Action.Execute', 
        title: 'Create New Record', 
        verb: 'show_create_form',
        data: { 
          action: 'show_create_form',
          ...baseData,
          recordTitle: originalRecordTitle,
          searchQuery: searchQuery || originalRecordTitle,
        } 
      },
      { type: 'Action.Execute', title: 'Deny Request', style: 'destructive', verb: 'deny_record', data: { action: 'deny_record', approvalId, recordUid: null, recordTitle: originalRecordTitle, requesterId, requesterEmail, requesterName, justification } },
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
        { type: 'TextBlock', text: 'Permission Level', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
        { type: 'Input.ChoiceSet', id: 'permission', value: 'view_only', choices: RECORD_PERMISSIONS },
        { type: 'TextBlock', text: 'Duration', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
        { type: 'Input.ChoiceSet', id: 'duration', value: '1h', choices: DURATION_OPTIONS }
      );
      
      card.actions = [
        { type: 'Action.Execute', title: 'Approve', style: 'positive', verb: 'approve_record', data: { action: 'approve_record', approvalId, recordUid: record.uid, recordTitle: record.title, requesterId, requesterEmail, requesterName, justification } },
        { type: 'Action.Execute', title: 'Create New Record', verb: 'show_create_form', data: { action: 'show_create_form', ...baseData, recordTitle: originalRecordTitle, searchQuery: searchQuery || originalRecordTitle } },
        { type: 'Action.Execute', title: 'Reset', verb: 'reset_record_card', data: { action: 'reset_record_card', ...baseData, recordTitle: originalRecordTitle } },
        { type: 'Action.Execute', title: 'Deny', style: 'destructive', verb: 'deny_record', data: { action: 'deny_record', approvalId, recordUid: record.uid, recordTitle: record.title, requesterId, requesterEmail, requesterName, justification } },
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
        { type: 'TextBlock', text: 'Permission Level', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
        { type: 'Input.ChoiceSet', id: 'permission', value: 'view_only', choices: RECORD_PERMISSIONS },
        { type: 'TextBlock', text: 'Duration', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
        { type: 'Input.ChoiceSet', id: 'duration', value: '1h', choices: DURATION_OPTIONS }
      );
      
      card.actions = [
        { type: 'Action.Execute', title: 'Approve Selected', style: 'positive', verb: 'approve_selected_record', data: { action: 'approve_selected_record', approvalId, requesterId, requesterEmail, requesterName, justification } },
        { type: 'Action.Execute', title: 'Create New Record', verb: 'show_create_form', data: { action: 'show_create_form', ...baseData, recordTitle: originalRecordTitle, searchQuery: searchQuery || originalRecordTitle } },
        { type: 'Action.Execute', title: '↩️ Reset', verb: 'reset_record_card', data: { action: 'reset_record_card', ...baseData, recordTitle: originalRecordTitle } },
        { type: 'Action.Execute', title: 'Deny Request', style: 'destructive', verb: 'deny_record', data: { action: 'deny_record', approvalId, recordUid: null, recordTitle: originalRecordTitle, requesterId, requesterEmail, requesterName, justification } },
      ];
    }
  }
  
  return card;
}

/**
 * Build a record approval card with status (for updating existing card)
 */
function buildRecordApprovalCardWithStatus({
  approvalId,
  requesterName,
  requesterEmail,
  recordTitle,
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
      { type: 'TextBlock', text: 'Record Access Request', weight: 'Bolder', size: 'ExtraLarge' },
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              { type: 'TextBlock', text: 'Requester:', weight: 'Bolder', size: 'Medium' },
              { type: 'TextBlock', text: requesterName || 'Unknown', color: 'Warning', size: 'Medium' },
              { type: 'TextBlock', text: 'Record:', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
              { type: 'TextBlock', text: recordTitle || 'Unknown', color: 'Warning', size: 'Medium' },
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
    
    if (recordTitle) {
      detailsItems.push({
        type: 'ColumnSet',
        columns: [
          { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: 'Access granted for:', weight: 'Bolder', size: 'Small' }] },
          { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: recordTitle, size: 'Small', color: 'Good' }] },
        ],
      });
    }
    
    if (permission) {
      detailsItems.push({
        type: 'ColumnSet',
        columns: [
          { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: 'Permission:', weight: 'Bolder', size: 'Small' }] },
          { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: formatPermissionLabel(permission), size: 'Small' }] },
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
 * Build a record confirmation card (after search/selection in task module)
 */
function buildRecordConfirmationCard({
  approvalId,
  requesterName,
  requesterId,
  requesterEmail,
  recordTitle,
  recordUid,
  justification,
  permission,
  duration,
}) {
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: [
      { type: 'TextBlock', text: 'Record Selected - Ready to Approve', weight: 'Bolder', size: 'Large', color: 'Good' },
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              { type: 'TextBlock', text: 'Requester:', weight: 'Bolder', size: 'Medium' },
              { type: 'TextBlock', text: requesterName || 'Unknown', color: 'Warning', size: 'Medium' },
              { type: 'TextBlock', text: 'Selected Record:', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
              { type: 'TextBlock', text: recordTitle || recordUid, color: 'Good', size: 'Medium', weight: 'Bolder' },
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
            { title: 'Permission:', value: formatPermissionLabel(permission) },
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
        verb: 'approve_record',
        data: { action: 'approve_record', approvalId, recordUid, recordTitle, requesterId, requesterEmail, requesterName, justification, permission, duration },
      },
      {
        type: 'Action.Execute',
        title: 'Deny',
        style: 'destructive',
        verb: 'deny_record',
        data: { action: 'deny_record', approvalId, recordUid, recordTitle, requesterId, requesterEmail, requesterName, justification },
      },
    ],
  };
}

/**
 * Build an inline record creation form card
 * Allows admin to create a new record directly on the approval card
 */
function buildRecordCreationCard({
  approvalId,
  requesterName,
  requesterId,
  requesterEmail,
  requesterAadObjectId,
  justification,
  identifier,
  originalRecordTitle,
  searchQuery,
  error,
  // Preserve form values on validation error
  recordTitle: prevRecordTitle,
  recordLogin: prevRecordLogin,
  recordPassword: prevRecordPassword,
  recordUrl: prevRecordUrl,
  recordNotes: prevRecordNotes,
}) {
  const headerElements = buildSearchCardHeader('Create New Record', requesterName || 'Unknown', approvalId || 'N/A', justification || '');
  
  const baseData = { 
    approvalId: approvalId || '', 
    identifier: identifier || '', 
    requesterId: requesterId || '', 
    requesterEmail: requesterEmail || '', 
    requesterAadObjectId: requesterAadObjectId || '', 
    requesterName: requesterName || '', 
    justification: justification || '', 
    originalRecordTitle: originalRecordTitle || '',
  };
  
  const bodyElements = [
    ...headerElements,
    { type: 'TextBlock', text: 'Create a new record to share with the requester:', wrap: true, spacing: 'Medium' },
  ];
  
  // Show validation error if present
  if (error) {
    bodyElements.push({
      type: 'TextBlock',
      text: `${error}`,
      wrap: true,
      color: 'Attention',
      weight: 'Bolder',
      spacing: 'Medium',
    });
  }
  
  // Form fields - preserve values on validation error
  bodyElements.push(
    // Title (required)
    { type: 'TextBlock', text: 'Title *', weight: 'Bolder', spacing: 'Medium' },
    { type: 'Input.Text', id: 'recordTitle', placeholder: 'Enter record title...', value: prevRecordTitle || searchQuery || originalRecordTitle || '' },
    
    // Login (required)
    { type: 'TextBlock', text: 'Login *', weight: 'Bolder', spacing: 'Small' },
    { type: 'Input.Text', id: 'recordLogin', placeholder: 'Enter username or email...', value: prevRecordLogin || '' },
    
    // Password (optional - will generate if empty)
    { type: 'TextBlock', text: 'Password (leave empty to auto-generate)', weight: 'Bolder', spacing: 'Small' },
    { type: 'Input.Text', id: 'recordPassword', placeholder: 'Enter password or leave empty...', style: 'password', value: prevRecordPassword || '' },
    
    // URL (optional)
    { type: 'TextBlock', text: 'URL (optional)', weight: 'Bolder', spacing: 'Small' },
    { type: 'Input.Text', id: 'recordUrl', placeholder: 'https://example.com', value: prevRecordUrl || '' },
    
    // Notes (optional)
    { type: 'TextBlock', text: 'Notes (optional)', weight: 'Bolder', spacing: 'Small' },
    { type: 'Input.Text', id: 'recordNotes', placeholder: 'Add any notes...', isMultiline: true, value: prevRecordNotes || '' },
    
    // Divider before self-destruct options
    { type: 'TextBlock', text: '─────────────────────────────', isSubtle: true, spacing: 'Medium' },
    
    // Self-destruct section header
    { type: 'TextBlock', text: 'Self-Destruct Options', weight: 'Bolder', size: 'Medium', spacing: 'Small' },
    { type: 'TextBlock', text: 'Enable this to automatically delete the record after a set time period.', size: 'Small', isSubtle: true, wrap: true },
    
    // Self-destruct toggle
    { type: 'Input.Toggle', id: 'selfDestruct', title: 'Enable Self-Destruct', value: 'false' },
    
    // Self-destruct duration dropdown
    { type: 'TextBlock', text: 'Delete After (only applies if self-destruct is enabled)', weight: 'Bolder', spacing: 'Small' },
    { type: 'Input.ChoiceSet', id: 'selfDestructDuration', value: '24h', choices: SELF_DESTRUCT_DURATION_OPTIONS },
    { type: 'TextBlock', text: 'This duration is ignored if self-destruct is not enabled above.', size: 'Small', isSubtle: true, wrap: true, color: 'Attention' },
  );
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: bodyElements,
    actions: [
      {
        type: 'Action.Execute',
        title: 'Create Record',
        style: 'positive',
        verb: 'submit_create_record',
        data: { action: 'submit_create_record', ...baseData },
      },
      {
        type: 'Action.Execute',
        title: 'Cancel',
        verb: 'cancel_create_form',
        data: { action: 'cancel_create_form', ...baseData, searchQuery },
      },
    ],
  };
}

/**
 * Build a card showing the newly created record with approval options
 */
function buildRecordCreatedCard({
  approvalId,
  requesterName,
  requesterId,
  requesterEmail,
  requesterAadObjectId,
  justification,
  identifier,
  originalRecordTitle,
  newRecordUid,
  newRecordTitle,
}) {
  const headerElements = buildSearchCardHeader('Record Access Request', requesterName || 'Unknown', approvalId || 'N/A', justification || '');
  
  // Ensure all values are defined
  const safeRecordTitle = newRecordTitle || 'New Record';
  const safeRecordUid = newRecordUid || 'Unknown';
  const safeRequesterId = requesterId || '';
  const safeRequesterEmail = requesterEmail || '';
  const safeRequesterName = requesterName || 'Unknown';
  const safeJustification = justification || '';
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'https://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: [
      ...headerElements,
      {
        type: 'Container',
        style: 'good',
        spacing: 'Medium',
        items: [
          { type: 'TextBlock', text: `Record Created: ${safeRecordTitle}`, wrap: true, weight: 'Bolder' },
          { type: 'TextBlock', text: `UID: ${safeRecordUid}`, size: 'Small', isSubtle: true },
        ],
      },
      { type: 'TextBlock', text: 'Now select permission and duration to grant access:', wrap: true, spacing: 'Medium' },
      { type: 'TextBlock', text: 'Permission Level', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
      { type: 'Input.ChoiceSet', id: 'permission', value: 'view_only', choices: RECORD_PERMISSIONS },
      { type: 'TextBlock', text: 'Duration', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
      { type: 'Input.ChoiceSet', id: 'duration', value: '1h', choices: DURATION_OPTIONS },
    ],
    actions: [
      {
        type: 'Action.Execute',
        title: 'Approve',
        style: 'positive',
        verb: 'approve_record',
        data: { 
          action: 'approve_record', 
          approvalId: approvalId || '', 
          recordUid: safeRecordUid, 
          recordTitle: safeRecordTitle, 
          requesterId: safeRequesterId, 
          requesterEmail: safeRequesterEmail, 
          requesterName: safeRequesterName, 
          justification: safeJustification,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny',
        style: 'destructive',
        verb: 'deny_record',
        data: { 
          action: 'deny_record', 
          approvalId: approvalId || '', 
          recordUid: safeRecordUid, 
          recordTitle: safeRecordTitle, 
          requesterId: safeRequesterId, 
          requesterEmail: safeRequesterEmail, 
          requesterName: safeRequesterName, 
          justification: safeJustification,
        },
      },
    ],
  };
}

module.exports = {
  buildRecordApprovalCard,
  buildRecordSearchResultsCard,
  buildRecordApprovalCardWithStatus,
  buildRecordConfirmationCard,
  buildRecordCreationCard,
  buildRecordCreatedCard,
};
