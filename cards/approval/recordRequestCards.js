/**
 * Record-related Adaptive Card builders
 * Cards for record access requests, search results, and confirmations
 */

const { RECORD_PERMISSIONS, DURATION_OPTIONS, SELF_DESTRUCT_DURATION_OPTIONS, DEFAULT_DURATION } = require('../constants');
const { 
  buildSearchCardHeader, 
  buildCreateSecretHeader,
  buildNoResultsSection, 
  formatPermissionLabel,
  getCurrentTimestamp,
} = require('../cardHelpers');
const { sanitizeHyperlinks } = require('../../utils/helpers');

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
  const requestedTime = getCurrentTimestamp();
  
  // Sanitize identifier and justification to prevent URL injection
  const safeIdentifier = sanitizeHyperlinks(identifier || recordTitle);
  const safeJustification = sanitizeHyperlinks(justification) || 'No justification provided';
  
  const card = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
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
  
  if (isUid) {
    // Add refresh property only when UID is resolved (not in search mode)
    // This prevents input field values from resetting during action execution
    // Refresh property enables auto-refresh for all users when message is edited
    // Omitting userIds enables auto-refresh for ALL users in channels with <60 members
    card.refresh = {
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
    };
    
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
      { type: 'Input.ChoiceSet', id: 'duration', value: DEFAULT_DURATION, choices: DURATION_OPTIONS },
      { type: 'TextBlock', text: 'Note: Can Share, Edit & Share, and Change Owner permissions grant permanent access (duration will be ignored).', wrap: true, isSubtle: true, size: 'Small', spacing: 'Small' }
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
      { type: 'TextBlock', text: 'Enter a search term and click Search to find the record.', wrap: true, isSubtle: true, size: 'Small' }
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
    version: '1.5',
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
        { type: 'Input.ChoiceSet', id: 'duration', value: '1h', choices: DURATION_OPTIONS },
        { type: 'TextBlock', text: 'Note: Can Share, Edit & Share, and Change Owner permissions grant permanent access (duration will be ignored).', wrap: true, isSubtle: true, size: 'Small', spacing: 'Small' }
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
        { type: 'Input.ChoiceSet', id: 'duration', value: '1h', choices: DURATION_OPTIONS },
        { type: 'TextBlock', text: 'Note: Can Share, Edit & Share, and Change Owner permissions grant permanent access (duration will be ignored).', wrap: true, isSubtle: true, size: 'Small', spacing: 'Small' }
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
  statusMessage,
  approverName,
  permission,
  duration,
  expiresAt,
  processedTime,
}) {
  let statusText;
  let containerStyle;
  
  if (status === 'approved') {
    statusText = 'APPROVED';
    containerStyle = 'good';
  } else if (status === 'owner') {
    statusText = statusMessage || 'USER ALREADY HAS FULL ACCESS (OWNER)';
    containerStyle = 'warning';
  } else {
    statusText = 'DENIED';
    containerStyle = 'attention';
  }
  
  const time = processedTime || getCurrentTimestamp();
  
  const card = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
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
        style: containerStyle,
        items: [
          { type: 'TextBlock', text: statusText, weight: 'Bolder', size: 'Large', horizontalAlignment: 'Center', wrap: true },
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
  } else if (status === 'owner') {
    const ownerDetailsItems = [
      {
        type: 'TextBlock',
        text: 'The selected user is already the owner of this record and has full permissions.',
        wrap: true,
        size: 'Small',
        spacing: 'Small',
      },
      {
        type: 'TextBlock',
        text: 'No action is needed - the user already has complete access.',
        wrap: true,
        size: 'Small',
        isSubtle: true,
      },
    ];
    
    if (requesterEmail) {
      ownerDetailsItems.unshift({
        type: 'ColumnSet',
        columns: [
          { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: 'Record Owner:', weight: 'Bolder', size: 'Small' }] },
          { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: requesterEmail, size: 'Small' }] },
        ],
      });
    }
    
    card.body.splice(2, 0, { type: 'Container', spacing: 'Medium', items: ownerDetailsItems });
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
    version: '1.5',
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
 * Result card after successful create-secret (self-service) submit
 */
function buildCreateSecretSuccessCard({
  recordTitle,
  recordUid,
}) {
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      { type: 'TextBlock', text: 'Secret created', weight: 'Bolder', size: 'Large', color: 'Good' },
      {
        type: 'FactSet',
        facts: [
          { title: 'Title', value: recordTitle || '—' },
          { title: 'Record UID', value: recordUid || '—' },
        ],
      },
      {
        type: 'TextBlock',
        text: 'Open the record in Keeper to view or edit credentials. Do not share passwords in Teams chat.',
        wrap: true,
        isSubtle: true,
        spacing: 'Medium',
      },
    ],
  };
}

/**
 * Build an inline record creation form card
 * Allows admin to create a new record directly on the approval card
 * @param {boolean} [createSecretFlow] - True for /keeper-create-secret (self-service form)
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
  createSecretFlow = false,
  /** @type {Array<{title:string,value:string}>|null} Only when createSecretFlow — from share-report (create-secret command) */
  sharedFolderChoices = null,
  sharedFoldersLoadError = null,
  /** When true (create-secret only): no share-report rows for this user — form and Create Secret are disabled; Cancel stays enabled */
  noSharedFoldersForUser = false,
  selectedTargetFolderUid = '_default_',
  /** @type {Array<{title:string,value:string}>|null} Subfolders inside the selected shared folder */
  subfolderChoices = null,
  subfolderLoadError = null,
  selectedSubfolderUid = '',
  // Preserve form values on validation error
  recordTitle: prevRecordTitle,
  recordLogin: prevRecordLogin,
  recordPassword: prevRecordPassword,
  recordUrl: prevRecordUrl,
  recordNotes: prevRecordNotes,
}) {
  const headerElements = createSecretFlow
    ? buildCreateSecretHeader(requesterName || 'Unknown', requesterEmail || '')
    : buildSearchCardHeader('Create New Record', requesterName || 'Unknown', approvalId || 'N/A', justification || '');
  
  const baseData = { 
    approvalId: approvalId || '', 
    identifier: identifier || '', 
    requesterId: requesterId || '', 
    requesterEmail: requesterEmail || '', 
    requesterAadObjectId: requesterAadObjectId || '', 
    requesterName: requesterName || '', 
    justification: justification || '', 
    originalRecordTitle: originalRecordTitle || '',
    createSecretFlow: !!createSecretFlow,
  };
  
  const introText = createSecretFlow
    ? 'Record title is required. Leave password empty to auto-generate one.'
    : 'Create a new record to share with the requester:';

  const secretFormLocked = createSecretFlow && noSharedFoldersForUser;
  const inputEnabled = !secretFormLocked;

  const bodyElements = [...headerElements];

  if (secretFormLocked) {
    bodyElements.push({
      type: 'TextBlock',
      text: "You don't have access to any shared folder. Please contact your Keeper administrator to request access.",
      wrap: true,
      color: 'Attention',
      weight: 'Bolder',
      spacing: 'Medium',
    });
  } else {
    bodyElements.push({ type: 'TextBlock', text: introText, wrap: true, spacing: 'Medium' });
  }
  
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

  let folderChoiceValue = '';
  if (
    createSecretFlow &&
    !secretFormLocked &&
    Array.isArray(sharedFolderChoices) &&
    sharedFolderChoices.length > 0
  ) {
    const allowed = new Set(sharedFolderChoices.map((c) => c.value));
    const raw = selectedTargetFolderUid != null ? String(selectedTargetFolderUid).trim() : '';
    folderChoiceValue = raw && allowed.has(raw) ? raw : '';
    bodyElements.push(
      { type: 'TextBlock', text: 'Save to folder *', weight: 'Bolder', spacing: 'Medium' },
    );
    bodyElements.push({
      type: 'Input.ChoiceSet',
      id: 'targetFolderUid',
      choices: sharedFolderChoices,
      value: folderChoiceValue,
      isRequired: true,
      errorMessage: 'Please select a shared folder',
      isEnabled: inputEnabled,
    });

    // Subfolder section
    const subfoldersRequested = Array.isArray(subfolderChoices);
    const subfoldersLoaded = subfoldersRequested && subfolderChoices.length > 0;
    const subfoldersEmpty = subfoldersRequested && subfolderChoices.length === 0 && !subfolderLoadError;
    if (subfoldersLoaded) {
      const selectedFolderName = (sharedFolderChoices || []).find((c) => c.value === folderChoiceValue)?.title || 'Selected folder';
      const rootChoice = { title: 'Parent Folder (' + selectedFolderName + ')', value: '_root_' };
      const allSubChoices = [rootChoice, ...subfolderChoices];
      const allowedSub = new Set(allSubChoices.map((c) => c.value));
      const rawSub = selectedSubfolderUid != null ? String(selectedSubfolderUid).trim() : '';
      const subValue = rawSub && allowedSub.has(rawSub) ? rawSub : '_root_';
      bodyElements.push(
        { type: 'TextBlock', text: 'Subfolder (optional)', weight: 'Bolder', spacing: 'Small' },
        {
          type: 'Input.ChoiceSet',
          id: 'targetSubfolderUid',
          choices: allSubChoices,
          value: subValue,
          isEnabled: inputEnabled,
        },
      );
    } else if (subfolderLoadError) {
      bodyElements.push({
        type: 'TextBlock',
        text: String(subfolderLoadError),
        wrap: true,
        color: 'Attention',
        size: 'Small',
        spacing: 'Small',
      });
    } else if (subfoldersEmpty) {
      bodyElements.push({
        type: 'TextBlock',
        text: 'No subfolders found under this folder.',
        wrap: true,
        isSubtle: true,
        size: 'Small',
        spacing: 'Small',
      });
    }

    bodyElements.push({
      type: 'ActionSet',
      spacing: 'Small',
      actions: [
        {
          type: 'Action.Execute',
          title: subfoldersLoaded ? '\uD83D\uDD04 Reload Subfolders' : '\uD83D\uDCC2 Load Subfolders',
          verb: 'load_subfolders',
          data: { action: 'load_subfolders', ...baseData },
        },
      ],
    });
  } else if (createSecretFlow && !secretFormLocked && sharedFoldersLoadError) {
    bodyElements.push({
      type: 'ColumnSet',
      spacing: 'Medium',
      columns: [
        {
          type: 'Column',
          width: 'stretch',
          verticalContentAlignment: 'Center',
          items: [
            {
              type: 'TextBlock',
              text: String(sharedFoldersLoadError),
              wrap: true,
              color: 'Attention',
            },
          ],
        },
        {
          type: 'Column',
          width: 'auto',
          verticalContentAlignment: 'Center',
          items: [
            {
              type: 'ActionSet',
              actions: [
                {
                  type: 'Action.Execute',
                  title: '\u{1F504} Retry',
                  verb: 'refresh_shared_folders',
                  data: { action: 'refresh_shared_folders', ...baseData },
                },
              ],
            },
          ],
        },
      ],
    });
  }
  
  // Form fields - preserve values on validation error
  bodyElements.push(
    { type: 'TextBlock', text: 'Title *', weight: 'Bolder', spacing: 'Medium' },
    {
      type: 'Input.Text',
      id: 'recordTitle',
      placeholder: 'Enter record title...',
      value: prevRecordTitle || searchQuery || originalRecordTitle || '',
      ...(createSecretFlow ? { isEnabled: inputEnabled } : {}),
    },
    { type: 'TextBlock', text: createSecretFlow ? 'Login' : 'Login *', weight: 'Bolder', spacing: 'Small' },
    {
      type: 'Input.Text',
      id: 'recordLogin',
      placeholder: createSecretFlow ? 'Username or email' : 'Enter username or email...',
      value: prevRecordLogin || '',
      ...(createSecretFlow ? { isEnabled: inputEnabled } : {}),
    },
    { type: 'TextBlock', text: 'Password', weight: 'Bolder', spacing: 'Small' },
    {
      type: 'Input.Text',
      id: 'recordPassword',
      placeholder: createSecretFlow ? 'Leave empty for no password' : 'Enter password or leave empty...',
      style: 'password',
      value: prevRecordPassword || '',
      ...(createSecretFlow ? { isEnabled: inputEnabled } : {}),
    },
    ...(createSecretFlow ? [
      {
        type: 'TextBlock',
        text: '\uD83D\uDD12 End-to-end encrypted. Your data is protected by Keeper\u2019s zero-knowledge architecture.',
        wrap: true,
        isSubtle: true,
        size: 'Small',
        spacing: 'None',
        color: 'Warning',
      },
      {
        type: 'Input.Toggle',
        id: 'autoGeneratePassword',
        title: 'Auto-generate a strong password',
        value: 'false',
        ...(inputEnabled ? {} : { isEnabled: false }),
      },
    ] : []),
    { type: 'TextBlock', text: createSecretFlow ? 'URL' : 'URL (optional)', weight: 'Bolder', spacing: 'Small' },
    {
      type: 'Input.Text',
      id: 'recordUrl',
      placeholder: 'https://example.com',
      value: prevRecordUrl || '',
      ...(createSecretFlow ? { isEnabled: inputEnabled } : {}),
    },
    { type: 'TextBlock', text: createSecretFlow ? 'Notes' : 'Notes (optional)', weight: 'Bolder', spacing: 'Small' },
    {
      type: 'Input.Text',
      id: 'recordNotes',
      placeholder: 'Add any notes...',
      isMultiline: true,
      value: prevRecordNotes || '',
      ...(createSecretFlow ? { isEnabled: inputEnabled } : {}),
    },
  );

  if (!createSecretFlow) {
    bodyElements.push(
      { type: 'TextBlock', text: '─────────────────────────────', isSubtle: true, spacing: 'Medium' },
      { type: 'TextBlock', text: 'Self-Destruct Options', weight: 'Bolder', size: 'Medium', spacing: 'Small' },
      { type: 'TextBlock', text: 'Enable this to automatically delete the record after a set time period.', size: 'Small', isSubtle: true, wrap: true },
      { type: 'Input.Toggle', id: 'selfDestruct', title: 'Enable Self-Destruct', value: 'false' },
      { type: 'TextBlock', text: 'Delete After (only applies if self-destruct is enabled)', weight: 'Bolder', spacing: 'Small' },
      { type: 'Input.ChoiceSet', id: 'selfDestructDuration', value: DEFAULT_DURATION, choices: SELF_DESTRUCT_DURATION_OPTIONS },
      { type: 'TextBlock', text: 'This duration is ignored if self-destruct is not enabled above.', size: 'Small', isSubtle: true, wrap: true, color: 'Attention' },
    );
  }
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: bodyElements,
    actions: [
      {
        type: 'Action.Execute',
        title: createSecretFlow ? 'Create Secret' : 'Create Record',
        style: 'positive',
        verb: 'submit_create_record',
        isEnabled: !secretFormLocked,
        data: { action: 'submit_create_record', ...baseData },
      },
      {
        type: 'Action.Execute',
        title: 'Cancel',
        verb: createSecretFlow ? 'cancel_create_secret' : 'cancel_create_form',
        isEnabled: true,
        associatedInputs: 'none',
        data: createSecretFlow
          ? { action: 'cancel_create_secret', ...baseData }
          : { action: 'cancel_create_form', ...baseData, searchQuery },
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
    version: '1.5',
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
      { type: 'TextBlock', text: 'Note: Can Share, Edit & Share, and Change Owner permissions grant permanent access (duration will be ignored).', wrap: true, isSubtle: true, size: 'Small', spacing: 'Small' },
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

/**
 * Build a card for when share invitation is sent (user doesn't have Keeper account yet)
 */
function buildRecordInvitationSentCard({
  approvalId,
  requesterName,
  requesterEmail,
  recordTitle,
  recordUid,
  justification,
  permission,
  approverName,
  processedTime,
}) {
  const time = processedTime || getCurrentTimestamp();
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      { 
        type: 'TextBlock', 
        text: 'Share Invitation Sent', 
        weight: 'Bolder', 
        size: 'ExtraLarge',
        color: 'Warning'
      },
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
              { type: 'TextBlock', text: recordTitle || recordUid || 'Unknown', color: 'Warning', size: 'Medium' },
            ],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              { type: 'TextBlock', text: 'Request ID:', weight: 'Bolder', size: 'Medium' },
              { type: 'TextBlock', text: approvalId || 'N/A', color: 'Warning', size: 'Medium' },
              { type: 'TextBlock', text: 'Permission:', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
              { type: 'TextBlock', text: formatPermissionLabel(permission) || 'View Only', size: 'Medium' },
            ],
          },
        ],
      },
      {
        type: 'Container',
        style: 'warning',
        spacing: 'Medium',
        items: [
          { 
            type: 'TextBlock', 
            text: 'INVITATION SENT', 
            weight: 'Bolder', 
            size: 'Large', 
            horizontalAlignment: 'Center' 
          },
        ],
      },
      {
        type: 'Container',
        spacing: 'Medium',
        items: [
          { 
            type: 'TextBlock', 
            text: 'Share invitation has been sent to the user\'s email.', 
            wrap: true,
            weight: 'Bolder'
          },
          { 
            type: 'TextBlock', 
            text: 'They must accept the invitation and create a Keeper account to access this record.',
            wrap: true,
            isSubtle: true
          },
        ],
      },
      {
        type: 'Container',
        spacing: 'Medium',
        style: 'emphasis',
        items: [
          { type: 'TextBlock', text: 'Next Steps for Requester:', weight: 'Bolder', size: 'Small' },
          { type: 'TextBlock', text: '1. Check email for the Keeper invitation', size: 'Small', wrap: true },
          { type: 'TextBlock', text: '2. Accept the invitation and create a Keeper account', size: 'Small', wrap: true },
          { type: 'TextBlock', text: '3. The record will be automatically shared with them', size: 'Small', wrap: true },
        ],
      },
      {
        type: 'Container',
        spacing: 'Medium',
        items: [
          { 
            type: 'TextBlock', 
            text: `Approved by: ${approverName || 'Unknown'} at ${time}`, 
            size: 'Small', 
            isSubtle: true,
            horizontalAlignment: 'Right'
          },
        ],
      },
    ],
    actions: [],
  };
}

/**
 * Build a card for when user already has access to the record
 */
function buildRecordAlreadyHasAccessCard({
  approvalId,
  requesterName,
  requesterEmail,
  recordTitle,
  recordUid,
  justification,
  currentPermission,
  currentPermissionLabel,
  approverName,
  processedTime,
}) {
  const time = processedTime || getCurrentTimestamp();
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      { 
        type: 'TextBlock', 
        text: 'User Already Has Access', 
        weight: 'Bolder', 
        size: 'ExtraLarge',
        color: 'Warning'
      },
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
              { type: 'TextBlock', text: recordTitle || recordUid || 'Unknown', color: 'Warning', size: 'Medium' },
            ],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              { type: 'TextBlock', text: 'Request ID:', weight: 'Bolder', size: 'Medium' },
              { type: 'TextBlock', text: approvalId || 'N/A', color: 'Warning', size: 'Medium' },
              { type: 'TextBlock', text: 'Current Permission:', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
              { type: 'TextBlock', text: currentPermissionLabel || currentPermission || 'Unknown', color: 'Good', size: 'Medium', weight: 'Bolder' },
            ],
          },
        ],
      },
      {
        type: 'Container',
        style: 'warning',
        spacing: 'Medium',
        items: [
          { 
            type: 'TextBlock', 
            text: 'NO ACTION NEEDED', 
            weight: 'Bolder', 
            size: 'Large', 
            horizontalAlignment: 'Center' 
          },
        ],
      },
      {
        type: 'Container',
        spacing: 'Medium',
        items: [
          { 
            type: 'TextBlock', 
            text: `The user "${requesterEmail || requesterName}" already has "${currentPermissionLabel || currentPermission}" access to this record.`, 
            wrap: true,
            weight: 'Bolder'
          },
          { 
            type: 'TextBlock', 
            text: 'No further action is required. The user can already access this record with the indicated permission level.',
            wrap: true,
            isSubtle: true
          },
        ],
      },
      {
        type: 'Container',
        spacing: 'Medium',
        items: [
          { 
            type: 'TextBlock', 
            text: `Checked by: ${approverName || 'Unknown'} at ${time}`, 
            size: 'Small', 
            isSubtle: true,
            horizontalAlignment: 'Right'
          },
        ],
      },
    ],
    actions: [],
  };
}

/**
 * Build a "Processing" card shown immediately when approval is clicked
 * This card is returned immediately to avoid Teams timeout, while the actual grant happens in background
 */
function buildRecordProcessingCard({
  approvalId,
  requesterName,
  requesterEmail,
  recordTitle,
  justification,
  permission,
  duration,
  approverName,
  processedTime,
}) {
  const time = processedTime || getCurrentTimestamp();
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
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
        style: 'emphasis',
        items: [
          { type: 'TextBlock', text: 'APPROVED - Processing...', weight: 'Bolder', size: 'Large', horizontalAlignment: 'Center', color: 'Good' },
          { type: 'TextBlock', text: 'Granting access to user. Please wait...', size: 'Small', horizontalAlignment: 'Center', isSubtle: true },
          { type: 'TextBlock', text: `By: ${approverName || 'Unknown'} at ${time}`, size: 'Small', horizontalAlignment: 'Center', isSubtle: true },
        ],
      },
      {
        type: 'ColumnSet',
        spacing: 'Medium',
        columns: [
          { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: 'Permission:', weight: 'Bolder', size: 'Small' }] },
          { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: formatPermissionLabel(permission), size: 'Small' }] },
        ],
      },
      {
        type: 'ColumnSet',
        columns: [
          { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: 'Duration:', weight: 'Bolder', size: 'Small' }] },
          { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: duration || '24h', size: 'Small' }] },
        ],
      },
      {
        type: 'ColumnSet',
        columns: [
          { type: 'Column', width: 'auto', items: [{ type: 'TextBlock', text: 'Granted To:', weight: 'Bolder', size: 'Small' }] },
          { type: 'Column', width: 'stretch', items: [{ type: 'TextBlock', text: requesterEmail || 'Unknown', size: 'Small' }] },
        ],
      },
    ],
    actions: [],
  };
}

module.exports = {
  buildRecordApprovalCard,
  buildRecordSearchResultsCard,
  buildRecordApprovalCardWithStatus,
  buildRecordConfirmationCard,
  buildRecordCreationCard,
  buildCreateSecretSuccessCard,
  buildRecordCreatedCard,
  buildRecordInvitationSentCard,
  buildRecordAlreadyHasAccessCard,
  buildRecordProcessingCard,
};
