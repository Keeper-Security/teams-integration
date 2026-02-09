/**
 * Adaptive Card builders for approval requests
 * 
 * These cards are displayed in the approvals channel when users request access
 * to records or folders. They include interactive elements for approvers to
 * select permissions and approve/deny requests.
 */

/**
 * Permission options for records
 */
const RECORD_PERMISSIONS = [
  { title: 'View Only', value: 'view_only' },
  { title: 'Can Edit', value: 'can_edit' },
  { title: 'Can Share (Permanent)', value: 'can_share' },
  { title: 'Edit & Share (Permanent)', value: 'edit_and_share' },
  { title: 'Change Owner (Permanent)', value: 'change_owner' },
];

/**
 * Permission options for folders
 */
const FOLDER_PERMISSIONS = [
  { title: 'No Permissions', value: 'no_permissions' },
  { title: 'Manage Users (Permanent)', value: 'manage_users' },
  { title: 'Manage Records', value: 'manage_records' },
  { title: 'Manage All (Permanent)', value: 'manage_all' },
];

/**
 * Duration options for time-limited access
 */
const DURATION_OPTIONS = [
  { title: '1 hour', value: '1h' },
  { title: '4 hours', value: '4h' },
  { title: '8 hours', value: '8h' },
  { title: '24 hours', value: '24h' },
  { title: '7 days', value: '7d' },
  { title: '30 days', value: '30d' },
  { title: 'Permanent', value: 'permanent' },
];

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
  // Format timestamp
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
          approvalId: approvalId,
          type: 'record',
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterName: requesterName,
          recordTitle: recordTitle,
          recordUid: recordUid,
          justification: justification,
          identifier: identifier,
          isUid: isUid,
        },
      },
      userIds: [], 
    },
    body: [
      // Header
      {
        type: 'TextBlock',
        text: 'Record Access Request',
        weight: 'Bolder',
        size: 'ExtraLarge',
      },
      // Two-column layout for details
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Requester:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: requesterName,
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Record:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: identifier || recordTitle,
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Requested:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: requestedTime,
                size: 'Medium',
              },
            ],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Request ID:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: approvalId,
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Justification:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: justification || 'No justification provided',
                wrap: true,
                size: 'Medium',
              },
            ],
          },
        ],
      },
    ],
    actions: [],
  };
  
  // Add content based on isUid
  if (isUid) {
    // Add Record Details section (like Slack) when UID is resolved
    if (recordTitle && recordTitle !== identifier) {
      card.body.push(
        {
          type: 'Container',
          separator: true,
          spacing: 'Medium',
          items: [
            {
              type: 'TextBlock',
              text: 'Record Details',
              weight: 'Bolder',
              size: 'Medium',
            },
            {
              type: 'FactSet',
              facts: [
                { title: 'Title:', value: recordTitle },
                { title: 'Type:', value: recordType || 'Record' },
              ],
            },
          ],
        }
      );
    }
    
    // Valid UID - show permission/duration selectors
    card.body.push(
      {
        type: 'TextBlock',
        text: 'Permission Level',
        weight: 'Bolder',
        size: 'Medium',
        spacing: 'Medium',
      },
      {
        type: 'Input.ChoiceSet',
        id: 'permission',
        value: 'view_only',
        choices: RECORD_PERMISSIONS,
      },
      {
        type: 'TextBlock',
        text: 'Duration',
        weight: 'Bolder',
        size: 'Medium',
        spacing: 'Medium',
      },
      {
        type: 'Input.ChoiceSet',
        id: 'duration',
        value: '24h',
        choices: DURATION_OPTIONS,
      }
    );
    
    card.actions = [
      {
        type: 'Action.Execute',
        title: 'Approve',
        style: 'positive',
        verb: 'approve_record',
        data: {
          action: 'approve_record',
          approvalId: approvalId,
          recordUid: recordUid,
          recordTitle: recordTitle,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterName: requesterName,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny',
        style: 'destructive',
        verb: 'deny_record',
        data: {
          action: 'deny_record',
          approvalId: approvalId,
          recordUid: recordUid,
          recordTitle: recordTitle,
          requesterId: requesterId,
          requesterName: requesterName,
        },
      },
    ];
  } else {
    // Description - show inline search input
    card.body.push(
      {
        type: 'TextBlock',
        text: '**Action Required:** Search for the correct record',
        wrap: true,
        size: 'Medium',
        spacing: 'Large',
      },
      {
        type: 'Input.Text',
        id: 'searchQuery',
        placeholder: 'Enter record name or UID to search...',
        value: identifier || recordTitle || '',
      },
      {
        type: 'TextBlock',
        text: 'Enter a search term and click Look Up to find the record.',
        wrap: true,
        isSubtle: true,
        size: 'Small',
      }
    );
    
    // Look Up and Deny buttons
    card.actions = [
      {
        type: 'Action.Execute',
        title: '🔍 Search',
        style: 'positive',
        verb: 'lookup_record',
        data: {
          action: 'lookup_record',
          approvalId: approvalId,
          identifier: identifier || recordTitle,
          recordTitle: recordTitle,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterAadObjectId: requesterAadObjectId,
          requesterName: requesterName,
          justification: justification,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny Request',
        style: 'destructive',
        verb: 'deny_record',
        data: {
          action: 'deny_record',
          approvalId: approvalId,
          recordUid: null,
          recordTitle: recordTitle,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterName: requesterName,
          justification: justification,
        },
      },
    ];
  }
  
  return card;
}

/**
 * Build a record search results card (inline search flow)
 * Shows found record with Approve/Deny/Reset buttons using Action.Execute
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
  // Search results - now supports array of records
  foundRecords, // Array of { uid, title }
  noResults = false,
  // Original card data for reset
  originalRecordTitle,
}) {
  const requestedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
  
  const card = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: [
      // Header
      {
        type: 'TextBlock',
        text: 'Record Access Request',
        weight: 'Bolder',
        size: 'ExtraLarge',
      },
      // Two-column layout for request details
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Requester:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: requesterName || 'Unknown',
                color: 'Warning',
                size: 'Medium',
              },
            ],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Request ID:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: approvalId || 'N/A',
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Justification:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: justification || 'No justification provided',
                wrap: true,
                size: 'Medium',
              },
            ],
          },
        ],
      },
    ],
    actions: [],
  };
  
  if (noResults) {
    // No results found
    card.body.push(
      {
        type: 'Container',
        style: 'attention',
        spacing: 'Medium',
        items: [
          {
            type: 'TextBlock',
            text: `No records found for "${searchQuery}"`,
            wrap: true,
            weight: 'Bolder',
          },
        ],
      },
      {
        type: 'TextBlock',
        text: 'Try a different search term:',
        wrap: true,
        spacing: 'Medium',
      },
      {
        type: 'Input.Text',
        id: 'searchQuery',
        placeholder: 'Enter record name or UID...',
        value: searchQuery || '',
      }
    );
    
    card.actions = [
      {
        type: 'Action.Execute',
        title: '🔍 Search',
        style: 'positive',
        verb: 'lookup_record',
        data: {
          action: 'lookup_record',
          approvalId: approvalId,
          identifier: identifier,
          recordTitle: originalRecordTitle,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterAadObjectId: requesterAadObjectId,
          requesterName: requesterName,
          justification: justification,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Reset',
        verb: 'reset_record_card',
        data: {
          action: 'reset_record_card',
          approvalId: approvalId,
          identifier: identifier,
          recordTitle: originalRecordTitle,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterAadObjectId: requesterAadObjectId,
          requesterName: requesterName,
          justification: justification,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny Request',
        style: 'destructive',
        verb: 'deny_record',
        data: {
          action: 'deny_record',
          approvalId: approvalId,
          recordUid: null,
          recordTitle: originalRecordTitle,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterName: requesterName,
          justification: justification,
        },
      },
    ];
  } else if (foundRecords && foundRecords.length > 0) {
    // Records found - show dropdown selection if multiple, or single record details
    const recordCount = foundRecords.length;
    
    if (recordCount === 1) {
      // Single record found - show direct approval
      const record = foundRecords[0];
      card.body.push(
        {
          type: 'Container',
          style: 'good',
          spacing: 'Medium',
          items: [
            {
              type: 'TextBlock',
              text: `Record Found: ${record.title}`,
              wrap: true,
              weight: 'Bolder',
            },
            {
              type: 'TextBlock',
              text: `UID: ${record.uid}`,
              size: 'Small',
              isSubtle: true,
            },
          ],
        },
        {
          type: 'TextBlock',
          text: 'Permission Level',
          weight: 'Bolder',
          size: 'Medium',
          spacing: 'Medium',
        },
        {
          type: 'Input.ChoiceSet',
          id: 'permission',
          value: 'view_only',
          choices: RECORD_PERMISSIONS,
        },
        {
          type: 'TextBlock',
          text: 'Duration',
          weight: 'Bolder',
          size: 'Medium',
          spacing: 'Medium',
        },
        {
          type: 'Input.ChoiceSet',
          id: 'duration',
          value: '1h',
          choices: DURATION_OPTIONS,
        }
      );
      
      card.actions = [
        {
          type: 'Action.Execute',
          title: 'Approve',
          style: 'positive',
          verb: 'approve_record',
          data: {
            action: 'approve_record',
            approvalId: approvalId,
            recordUid: record.uid,
            recordTitle: record.title,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterName: requesterName,
            justification: justification,
          },
        },
        {
          type: 'Action.Execute',
          title: 'Reset',
          verb: 'reset_record_card',
          data: {
            action: 'reset_record_card',
            approvalId: approvalId,
            identifier: identifier,
            recordTitle: originalRecordTitle,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterAadObjectId: requesterAadObjectId,
            requesterName: requesterName,
            justification: justification,
          },
        },
        {
          type: 'Action.Execute',
          title: 'Deny',
          style: 'destructive',
          verb: 'deny_record',
          data: {
            action: 'deny_record',
            approvalId: approvalId,
            recordUid: record.uid,
            recordTitle: record.title,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterName: requesterName,
            justification: justification,
          },
        },
      ];
    } else {
      // Multiple records found - show dropdown selection
      const recordChoices = foundRecords.map(r => ({
        title: `${r.title} (${r.uid.substring(0, 8)}...)`,
        value: JSON.stringify({ uid: r.uid, title: r.title }),
      }));
      
      card.body.push(
        {
          type: 'Container',
          style: 'good',
          spacing: 'Medium',
          items: [
            {
              type: 'TextBlock',
              text: `${recordCount} Records Found`,
              wrap: true,
              weight: 'Bolder',
            },
            {
              type: 'TextBlock',
              text: 'Select the correct record from the list below:',
              size: 'Small',
              isSubtle: true,
            },
          ],
        },
        {
          type: 'TextBlock',
          text: 'Select Record',
          weight: 'Bolder',
          size: 'Medium',
          spacing: 'Medium',
        },
        {
          type: 'Input.ChoiceSet',
          id: 'selectedRecord',
          value: recordChoices[0].value,
          choices: recordChoices,
          style: 'expanded',
        },
        {
          type: 'TextBlock',
          text: 'Permission Level',
          weight: 'Bolder',
          size: 'Medium',
          spacing: 'Medium',
        },
        {
          type: 'Input.ChoiceSet',
          id: 'permission',
          value: 'view_only',
          choices: RECORD_PERMISSIONS,
        },
        {
          type: 'TextBlock',
          text: 'Duration',
          weight: 'Bolder',
          size: 'Medium',
          spacing: 'Medium',
        },
        {
          type: 'Input.ChoiceSet',
          id: 'duration',
          value: '1h',
          choices: DURATION_OPTIONS,
        }
      );
      
      // For multiple records, we pass the full list and let the handler pick based on selection
      card.actions = [
        {
          type: 'Action.Execute',
          title: 'Approve Selected',
          style: 'positive',
          verb: 'approve_selected_record',
          data: {
            action: 'approve_selected_record',
            approvalId: approvalId,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterName: requesterName,
            justification: justification,
          },
        },
        {
          type: 'Action.Execute',
          title: '↩️ Reset',
          verb: 'reset_record_card',
          data: {
            action: 'reset_record_card',
            approvalId: approvalId,
            identifier: identifier,
            recordTitle: originalRecordTitle,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterAadObjectId: requesterAadObjectId,
            requesterName: requesterName,
            justification: justification,
          },
        },
        {
          type: 'Action.Execute',
          title: 'Deny Request',
          style: 'destructive',
          verb: 'deny_record',
          data: {
            action: 'deny_record',
            approvalId: approvalId,
            recordUid: null,
            recordTitle: originalRecordTitle,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterName: requesterName,
            justification: justification,
          },
        },
      ];
    }
  }
  
  return card;
}

/**
 * Build an Adaptive Card for folder access approval request
 * Slack-style layout with two columns (matching record card style)
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
  // Format timestamp
  const requestedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
  
  const card = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    // Auto-refresh: when card is viewed, Teams will invoke this action to get updated card
    refresh: {
      action: {
        type: 'Action.Execute',
        verb: 'refreshApprovalCard',
        data: {
          approvalId: approvalId,
          type: 'folder',
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterName: requesterName,
          folderName: folderName,
          folderUid: folderUid,
          justification: justification,
          identifier: identifier,
          isUid: isUid,
        },
      },
      userIds: [], // Empty array = refresh for ALL users who view the card
    },
    body: [
      // Header
      {
        type: 'TextBlock',
        text: 'Folder Access Request',
        weight: 'Bolder',
        size: 'ExtraLarge',
      },
      // Two-column layout for details
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Requester:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: requesterName,
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Folder:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: identifier || folderName,
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Requested:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: requestedTime,
                size: 'Medium',
              },
            ],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Request ID:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: approvalId,
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Justification:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: justification || 'No justification provided',
                wrap: true,
                size: 'Medium',
              },
            ],
          },
        ],
      },
    ],
    actions: [],
  };
  
  // Add content based on isUid
  if (isUid) {
    // Add Folder Details section (like Slack) when UID is resolved
    if (folderName && folderName !== identifier) {
      card.body.push(
        {
          type: 'Container',
          separator: true,
          spacing: 'Medium',
          items: [
            {
              type: 'TextBlock',
              text: 'Folder Details',
              weight: 'Bolder',
              size: 'Medium',
            },
            {
              type: 'FactSet',
              facts: [
                { title: 'Title:', value: folderName },
                { title: 'Type:', value: folderType || 'Shared Folder' },
              ],
            },
          ],
        }
      );
    }
    
    // Valid UID - show permission/duration selectors
    card.body.push(
      {
        type: 'TextBlock',
        text: 'Permission Level',
        weight: 'Bolder',
        size: 'Medium',
        spacing: 'Medium',
      },
      {
        type: 'Input.ChoiceSet',
        id: 'permission',
        value: 'no_permissions',
        choices: FOLDER_PERMISSIONS,
      },
      {
        type: 'TextBlock',
        text: 'Duration',
        weight: 'Bolder',
        size: 'Medium',
        spacing: 'Medium',
      },
      {
        type: 'Input.ChoiceSet',
        id: 'duration',
        value: '24h',
        choices: DURATION_OPTIONS,
      }
    );
    
    card.actions = [
      {
        type: 'Action.Execute',
        title: 'Approve',
        style: 'positive',
        verb: 'approve_folder',
        data: {
          action: 'approve_folder',
          approvalId: approvalId,
          folderUid: folderUid,
          folderName: folderName,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterName: requesterName,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny',
        style: 'destructive',
        verb: 'deny_folder',
        data: {
          action: 'deny_folder',
          approvalId: approvalId,
          folderUid: folderUid,
          folderName: folderName,
          requesterId: requesterId,
          requesterName: requesterName,
        },
      },
    ];
  } else {
    // Description - show inline search input
    card.body.push(
      {
        type: 'TextBlock',
        text: '**Action Required:** Search for the correct folder',
        wrap: true,
        size: 'Medium',
        spacing: 'Large',
      },
      {
        type: 'Input.Text',
        id: 'searchQuery',
        placeholder: 'Enter folder name or UID to search...',
        value: identifier || folderName || '',
      },
      {
        type: 'TextBlock',
        text: 'Enter a search term and click Look Up to find the folder.',
        wrap: true,
        isSubtle: true,
        size: 'Small',
      }
    );
    
    // Look Up and Deny buttons
    card.actions = [
      {
        type: 'Action.Execute',
        title: '🔍 Look Up',
        style: 'positive',
        verb: 'lookup_folder',
        data: {
          action: 'lookup_folder',
          approvalId: approvalId,
          identifier: identifier || folderName,
          folderName: folderName,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterAadObjectId: requesterAadObjectId,
          requesterName: requesterName,
          justification: justification,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny Request',
        style: 'destructive',
        verb: 'deny_folder',
        data: {
          action: 'deny_folder',
          approvalId: approvalId,
          folderUid: null,
          folderName: folderName,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterName: requesterName,
          justification: justification,
        },
      },
    ];
  }
  
  return card;
}

/**
 * Build a folder search results card (inline search flow)
 * Shows found folder(s) with Approve/Deny/Reset buttons using Action.Execute
 * Supports multiple results with dropdown selection
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
  // Search results - now supports array of folders
  foundFolders, // Array of { uid, name }
  noResults = false,
  // Original card data for reset
  originalFolderName,
}) {
  const requestedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
  
  const card = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: [
      // Header
      {
        type: 'TextBlock',
        text: 'Folder Access Request',
        weight: 'Bolder',
        size: 'ExtraLarge',
      },
      // Two-column layout for request details
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Requester:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: requesterName || 'Unknown',
                color: 'Warning',
                size: 'Medium',
              },
            ],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Request ID:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: approvalId || 'N/A',
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Justification:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: justification || 'No justification provided',
                wrap: true,
                size: 'Medium',
              },
            ],
          },
        ],
      },
    ],
    actions: [],
  };
  
  if (noResults) {
    // No results found
    card.body.push(
      {
        type: 'Container',
        style: 'attention',
        spacing: 'Medium',
        items: [
          {
            type: 'TextBlock',
            text: `No folders found for "${searchQuery}"`,
            wrap: true,
            weight: 'Bolder',
          },
        ],
      },
      {
        type: 'TextBlock',
        text: 'Try a different search term:',
        wrap: true,
        spacing: 'Medium',
      },
      {
        type: 'Input.Text',
        id: 'searchQuery',
        placeholder: 'Enter folder name or UID...',
        value: searchQuery || '',
      }
    );
    
    card.actions = [
      {
        type: 'Action.Execute',
        title: '🔍 Look Up',
        style: 'positive',
        verb: 'lookup_folder',
        data: {
          action: 'lookup_folder',
          approvalId: approvalId,
          identifier: identifier,
          folderName: originalFolderName,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterAadObjectId: requesterAadObjectId,
          requesterName: requesterName,
          justification: justification,
        },
      },
      {
        type: 'Action.Execute',
        title: '↩️ Reset',
        verb: 'reset_folder_card',
        data: {
          action: 'reset_folder_card',
          approvalId: approvalId,
          identifier: identifier,
          folderName: originalFolderName,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterAadObjectId: requesterAadObjectId,
          requesterName: requesterName,
          justification: justification,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny Request',
        style: 'destructive',
        verb: 'deny_folder',
        data: {
          action: 'deny_folder',
          approvalId: approvalId,
          folderUid: null,
          folderName: originalFolderName,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterName: requesterName,
          justification: justification,
        },
      },
    ];
  } else if (foundFolders && foundFolders.length > 0) {
    // Folders found - show dropdown selection if multiple, or single folder details
    const folderCount = foundFolders.length;
    
    if (folderCount === 1) {
      // Single folder found - show direct approval
      const folder = foundFolders[0];
      card.body.push(
        {
          type: 'Container',
          style: 'good',
          spacing: 'Medium',
          items: [
            {
              type: 'TextBlock',
              text: `Folder Found: ${folder.name}`,
              wrap: true,
              weight: 'Bolder',
            },
            {
              type: 'TextBlock',
              text: `UID: ${folder.uid}`,
              size: 'Small',
              isSubtle: true,
            },
          ],
        },
        {
          type: 'TextBlock',
          text: 'Permission Level',
          weight: 'Bolder',
          size: 'Medium',
          spacing: 'Medium',
        },
        {
          type: 'Input.ChoiceSet',
          id: 'permission',
          value: 'no_permissions',
          choices: FOLDER_PERMISSIONS,
        },
        {
          type: 'TextBlock',
          text: 'Duration',
          weight: 'Bolder',
          size: 'Medium',
          spacing: 'Medium',
        },
        {
          type: 'Input.ChoiceSet',
          id: 'duration',
          value: '1h',
          choices: DURATION_OPTIONS,
        }
      );
      
      card.actions = [
        {
          type: 'Action.Execute',
          title: 'Approve',
          style: 'positive',
          verb: 'approve_folder',
          data: {
            action: 'approve_folder',
            approvalId: approvalId,
            folderUid: folder.uid,
            folderName: folder.name,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterName: requesterName,
            justification: justification,
          },
        },
        {
          type: 'Action.Execute',
          title: '↩️ Reset',
          verb: 'reset_folder_card',
          data: {
            action: 'reset_folder_card',
            approvalId: approvalId,
            identifier: identifier,
            folderName: originalFolderName,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterAadObjectId: requesterAadObjectId,
            requesterName: requesterName,
            justification: justification,
          },
        },
        {
          type: 'Action.Execute',
          title: 'Deny',
          style: 'destructive',
          verb: 'deny_folder',
          data: {
            action: 'deny_folder',
            approvalId: approvalId,
            folderUid: folder.uid,
            folderName: folder.name,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterName: requesterName,
            justification: justification,
          },
        },
      ];
    } else {
      // Multiple folders found - show dropdown selection
      const folderChoices = foundFolders.map(f => ({
        title: `${f.name} (${f.uid.substring(0, 8)}...)`,
        value: JSON.stringify({ uid: f.uid, name: f.name }),
      }));
      
      card.body.push(
        {
          type: 'Container',
          style: 'good',
          spacing: 'Medium',
          items: [
            {
              type: 'TextBlock',
              text: `${folderCount} Folders Found`,
              wrap: true,
              weight: 'Bolder',
            },
            {
              type: 'TextBlock',
              text: 'Select the correct folder from the list below:',
              size: 'Small',
              isSubtle: true,
            },
          ],
        },
        {
          type: 'TextBlock',
          text: 'Select Folder',
          weight: 'Bolder',
          size: 'Medium',
          spacing: 'Medium',
        },
        {
          type: 'Input.ChoiceSet',
          id: 'selectedFolder',
          value: folderChoices[0].value,
          choices: folderChoices,
          style: 'expanded',
        },
        {
          type: 'TextBlock',
          text: 'Permission Level',
          weight: 'Bolder',
          size: 'Medium',
          spacing: 'Medium',
        },
        {
          type: 'Input.ChoiceSet',
          id: 'permission',
          value: 'no_permissions',
          choices: FOLDER_PERMISSIONS,
        },
        {
          type: 'TextBlock',
          text: 'Duration',
          weight: 'Bolder',
          size: 'Medium',
          spacing: 'Medium',
        },
        {
          type: 'Input.ChoiceSet',
          id: 'duration',
          value: '1h',
          choices: DURATION_OPTIONS,
        }
      );
      
      // For multiple folders, we pass the full list and let the handler pick based on selection
      card.actions = [
        {
          type: 'Action.Execute',
          title: 'Approve Selected',
          style: 'positive',
          verb: 'approve_selected_folder',
          data: {
            action: 'approve_selected_folder',
            approvalId: approvalId,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterName: requesterName,
            justification: justification,
          },
        },
        {
          type: 'Action.Execute',
          title: 'Reset',
          verb: 'reset_folder_card',
          data: {
            action: 'reset_folder_card',
            approvalId: approvalId,
            identifier: identifier,
            folderName: originalFolderName,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterAadObjectId: requesterAadObjectId,
            requesterName: requesterName,
            justification: justification,
          },
        },
        {
          type: 'Action.Execute',
          title: 'Deny Request',
          style: 'destructive',
          verb: 'deny_folder',
          data: {
            action: 'deny_folder',
            approvalId: approvalId,
            folderUid: null,
            folderName: originalFolderName,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterName: requesterName,
            justification: justification,
          },
        },
      ];
    }
  }
  
  return card;
}

/**
 * Build a search results card for one-time share
 * Shows search results and allows admin to select a record to share
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
  foundRecords, // Array of { uid, title }
  noResults = false,
  originalRecordTitle,
}) {
  const requestedTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
  
  const SHARE_DURATION_OPTIONS = [
    { title: '1 hour', value: '1h' },
    { title: '4 hours', value: '4h' },
    { title: '24 hours', value: '24h' },
    { title: '7 days', value: '7d' },
  ];
  
  const card = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: [
      // Header
      {
        type: 'TextBlock',
        text: 'One-Time Share Request',
        weight: 'Bolder',
        size: 'ExtraLarge',
      },
      // Two-column layout for request details
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Requester:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: requesterName || 'Unknown',
                color: 'Warning',
                size: 'Medium',
              },
            ],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Request ID:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: approvalId || 'N/A',
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Justification:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: justification || 'No justification provided',
                wrap: true,
                size: 'Medium',
              },
            ],
          },
        ],
      },
    ],
    actions: [],
  };
  
  if (noResults) {
    // No results found
    card.body.push(
      {
        type: 'Container',
        style: 'attention',
        spacing: 'Medium',
        items: [
          {
            type: 'TextBlock',
            text: `No records found for "${searchQuery}"`,
            wrap: true,
            weight: 'Bolder',
          },
        ],
      },
      {
        type: 'TextBlock',
        text: 'Try a different search term:',
        wrap: true,
        spacing: 'Medium',
      },
      {
        type: 'Input.Text',
        id: 'searchQuery',
        placeholder: 'Enter record name or UID...',
        value: searchQuery || '',
      }
    );
    
    card.actions = [
      {
        type: 'Action.Execute',
        title: '🔍 Search',
        style: 'positive',
        verb: 'lookup_share',
        data: {
          action: 'lookup_share',
          approvalId: approvalId,
          identifier: identifier,
          recordTitle: originalRecordTitle,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterAadObjectId: requesterAadObjectId,
          requesterName: requesterName,
          justification: justification,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Reset',
        verb: 'reset_share_card',
        data: {
          action: 'reset_share_card',
          approvalId: approvalId,
          identifier: identifier,
          recordTitle: originalRecordTitle,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterAadObjectId: requesterAadObjectId,
          requesterName: requesterName,
          justification: justification,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny Request',
        style: 'destructive',
        verb: 'deny_share',
        data: {
          action: 'deny_share',
          approvalId: approvalId,
          recordUid: null,
          recordTitle: originalRecordTitle,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterName: requesterName,
          justification: justification,
        },
      },
    ];
  } else if (foundRecords && foundRecords.length > 0) {
    // Records found - show dropdown selection if multiple, or single record details
    const recordCount = foundRecords.length;
    
    if (recordCount === 1) {
      // Single record found - show direct approval
      const record = foundRecords[0];
      card.body.push(
        {
          type: 'Container',
          style: 'good',
          spacing: 'Medium',
          items: [
            {
              type: 'TextBlock',
              text: `Record Found: ${record.title}`,
              wrap: true,
              weight: 'Bolder',
            },
            {
              type: 'TextBlock',
              text: `UID: ${record.uid}`,
              size: 'Small',
              isSubtle: true,
            },
          ],
        },
        {
          type: 'TextBlock',
          text: 'Share Duration',
          weight: 'Bolder',
          size: 'Medium',
          spacing: 'Medium',
        },
        {
          type: 'Input.ChoiceSet',
          id: 'duration',
          value: '24h',
          choices: SHARE_DURATION_OPTIONS,
        },
        {
          type: 'Input.Toggle',
          id: 'editable',
          title: 'Allow editing (recipient can modify the record)',
          value: 'false',
        }
      );
      
      card.actions = [
        {
          type: 'Action.Execute',
          title: 'Create Share',
          style: 'positive',
          verb: 'approve_share',
          data: {
            action: 'approve_share',
            approvalId: approvalId,
            recordUid: record.uid,
            recordTitle: record.title,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterName: requesterName,
            justification: justification,
          },
        },
        {
          type: 'Action.Execute',
          title: 'Reset',
          verb: 'reset_share_card',
          data: {
            action: 'reset_share_card',
            approvalId: approvalId,
            identifier: identifier,
            recordTitle: originalRecordTitle,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterAadObjectId: requesterAadObjectId,
            requesterName: requesterName,
            justification: justification,
          },
        },
        {
          type: 'Action.Execute',
          title: 'Deny',
          style: 'destructive',
          verb: 'deny_share',
          data: {
            action: 'deny_share',
            approvalId: approvalId,
            recordUid: record.uid,
            recordTitle: record.title,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterName: requesterName,
            justification: justification,
          },
        },
      ];
    } else {
      // Multiple records found - show dropdown selection
      const recordChoices = foundRecords.map(r => ({
        title: `${r.title} (${r.uid.substring(0, 8)}...)`,
        value: JSON.stringify({ uid: r.uid, title: r.title }),
      }));
      
      card.body.push(
        {
          type: 'Container',
          style: 'good',
          spacing: 'Medium',
          items: [
            {
              type: 'TextBlock',
              text: `${recordCount} Records Found`,
              wrap: true,
              weight: 'Bolder',
            },
            {
              type: 'TextBlock',
              text: 'Select the correct record from the list below:',
              size: 'Small',
              isSubtle: true,
            },
          ],
        },
        {
          type: 'TextBlock',
          text: 'Select Record',
          weight: 'Bolder',
          size: 'Medium',
          spacing: 'Medium',
        },
        {
          type: 'Input.ChoiceSet',
          id: 'selectedRecord',
          value: recordChoices[0].value,
          choices: recordChoices,
          style: 'expanded',
        },
        {
          type: 'TextBlock',
          text: 'Share Duration',
          weight: 'Bolder',
          size: 'Medium',
          spacing: 'Medium',
        },
        {
          type: 'Input.ChoiceSet',
          id: 'duration',
          value: '24h',
          choices: SHARE_DURATION_OPTIONS,
        },
        {
          type: 'Input.Toggle',
          id: 'editable',
          title: 'Allow editing (recipient can modify the record)',
          value: 'false',
        }
      );
      
      card.actions = [
        {
          type: 'Action.Execute',
          title: 'Create Share',
          style: 'positive',
          verb: 'approve_selected_share',
          data: {
            action: 'approve_selected_share',
            approvalId: approvalId,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterName: requesterName,
            justification: justification,
          },
        },
        {
          type: 'Action.Execute',
          title: 'Reset',
          verb: 'reset_share_card',
          data: {
            action: 'reset_share_card',
            approvalId: approvalId,
            identifier: identifier,
            recordTitle: originalRecordTitle,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterAadObjectId: requesterAadObjectId,
            requesterName: requesterName,
            justification: justification,
          },
        },
        {
          type: 'Action.Execute',
          title: 'Deny Request',
          style: 'destructive',
          verb: 'deny_share',
          data: {
            action: 'deny_share',
            approvalId: approvalId,
            recordUid: null,
            recordTitle: originalRecordTitle,
            requesterId: requesterId,
            requesterEmail: requesterEmail,
            requesterName: requesterName,
            justification: justification,
          },
        },
      ];
    }
  }
  
  return card;
}

/**
 * Build an Adaptive Card for one-time share approval request
 * Supports both UID-based (direct approval) and description-based (search flow)
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
  
  const card = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      // Header
      {
        type: 'TextBlock',
        text: 'One-Time Share Request',
        weight: 'Bolder',
        size: 'ExtraLarge',
      },
      // Two-column layout for details
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Requester:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: requesterName,
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Record:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: identifier || recordTitle,
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Requested:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: requestedTime,
                size: 'Medium',
              },
            ],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Request ID:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: approvalId,
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Justification:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: justification || 'No justification provided',
                wrap: true,
                size: 'Medium',
              },
            ],
          },
        ],
      },
    ],
    actions: [],
  };
  
  // Add content based on isUid
  if (isUid && recordUid) {
    // Valid UID - show record details and duration/editable options
    if (recordTitle && recordTitle !== identifier) {
      card.body.push({
        type: 'Container',
        separator: true,
        spacing: 'Medium',
        items: [
          {
            type: 'TextBlock',
            text: 'Record Details',
            weight: 'Bolder',
            size: 'Medium',
          },
          {
            type: 'FactSet',
            facts: [
              { title: 'Title:', value: recordTitle },
              { title: 'UID:', value: recordUid },
            ],
          },
        ],
      });
    }
    
    // Duration and editable options
    card.body.push(
      {
        type: 'TextBlock',
        text: 'Share Duration',
        weight: 'Bolder',
        size: 'Medium',
        spacing: 'Medium',
      },
      {
        type: 'Input.ChoiceSet',
        id: 'duration',
        value: '24h',
        choices: [
          { title: '1 hour', value: '1h' },
          { title: '4 hours', value: '4h' },
          { title: '24 hours', value: '24h' },
          { title: '7 days', value: '7d' },
        ],
      },
      {
        type: 'Input.Toggle',
        id: 'editable',
        title: 'Allow editing (recipient can modify the record)',
        value: 'false',
      }
    );
    
    card.actions = [
      {
        type: 'Action.Execute',
        title: 'Approve',
        style: 'positive',
        verb: 'approve_share',
        data: {
          action: 'approve_share',
          approvalId: approvalId,
          recordUid: recordUid,
          recordTitle: recordTitle,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterName: requesterName,
          justification: justification,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny',
        style: 'destructive',
        verb: 'deny_share',
        data: {
          action: 'deny_share',
          approvalId: approvalId,
          recordUid: recordUid,
          recordTitle: recordTitle,
          requesterId: requesterId,
          requesterName: requesterName,
        },
      },
    ];
  } else {
    // Description-based - show inline search input (same as record flow)
    card.body.push(
      {
        type: 'TextBlock',
        text: '**Action Required:** Search for the correct record to share',
        wrap: true,
        size: 'Medium',
        spacing: 'Large',
      },
      {
        type: 'Input.Text',
        id: 'searchQuery',
        placeholder: 'Enter record name or UID to search...',
        value: identifier || recordTitle || '',
      },
      {
        type: 'TextBlock',
        text: 'Enter a search term and click Search to find the record.',
        wrap: true,
        isSubtle: true,
        size: 'Small',
      }
    );
    
    // Search and Deny buttons
    card.actions = [
      {
        type: 'Action.Execute',
        title: '🔍 Search',
        style: 'positive',
        verb: 'lookup_share',
        data: {
          action: 'lookup_share',
          approvalId: approvalId,
          identifier: identifier || recordTitle,
          recordTitle: recordTitle,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterAadObjectId: requesterAadObjectId,
          requesterName: requesterName,
          justification: justification,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny Request',
        style: 'destructive',
        verb: 'deny_share',
        data: {
          action: 'deny_share',
          approvalId: approvalId,
          recordUid: null,
          recordTitle: identifier || recordTitle,
          requesterId: requesterId,
          requesterName: requesterName,
        },
      },
    ];
  }
  
  return card;
}

/**
 * Build a one-time share approval card with status (approved/denied)
 * Shows the same card but with status and no action buttons
 */
function buildOneTimeShareApprovalCardWithStatus({
  approvalId,
  requesterName,
  requesterEmail,
  recordTitle,
  recordUid,
  justification,
  status, // 'approved' or 'denied'
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
      // Header with status
      {
        type: 'Container',
        style: status === 'approved' ? 'good' : 'attention',
        items: [
          {
            type: 'TextBlock',
            text: 'One-Time Share Request',
            weight: 'Bolder',
            size: 'Large',
          },
          {
            type: 'TextBlock',
            text: statusText,
            weight: 'Bolder',
            color: statusColor,
          },
        ],
      },
      // Request details
      {
        type: 'Container',
        items: [
          {
            type: 'FactSet',
            facts: [
              { title: 'Request ID', value: approvalId },
              { title: 'Requester', value: requesterName },
              { title: 'Record', value: recordTitle },
            ],
          },
        ],
      },
      // Justification
      {
        type: 'Container',
        items: [
          {
            type: 'TextBlock',
            text: 'Justification',
            weight: 'Bolder',
            size: 'Medium',
          },
          {
            type: 'TextBlock',
            text: justification || 'No justification provided',
            wrap: true,
          },
        ],
      },
    ],
  };
  
  // Add approval details if approved
  if (status === 'approved') {
    card.body.push({
      type: 'Container',
      separator: true,
      items: [
        {
          type: 'TextBlock',
          text: 'Share Details',
          weight: 'Bolder',
          size: 'Medium',
          color: 'Good',
        },
        {
          type: 'FactSet',
          facts: [
            { title: 'Reviewed by', value: approverName },
            { title: 'Duration', value: duration || '24h' },
            { title: 'Editable', value: editable ? 'Yes' : 'No' },
            { title: 'Expires', value: expiresAt || 'N/A' },
            { title: 'Processed', value: time },
          ],
        },
        {
          type: 'TextBlock',
          text: '💡 Share link has been sent to the requester via DM.',
          wrap: true,
          isSubtle: true,
        },
      ],
    });
  } else {
    // Denied status
    card.body.push({
      type: 'Container',
      separator: true,
      items: [
        {
          type: 'FactSet',
          facts: [
            { title: 'Denied by', value: approverName },
            { title: 'Time', value: time },
          ],
        },
      ],
    });
  }
  
  return card;
}

/**
 * Build a record approval card with status (for updating existing card)
 * Shows the same card but with APPROVED/DENIED status and no action buttons
 */
function buildRecordApprovalCardWithStatus({
  approvalId,
  requesterName,
  requesterEmail,
  recordTitle,
  justification,
  status, // 'approved' or 'denied'
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
      // Header
      {
        type: 'TextBlock',
        text: 'Record Access Request',
        weight: 'Bolder',
        size: 'ExtraLarge',
      },
      // Two-column layout for details
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Requester:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: requesterName || 'Unknown',
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Record:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: recordTitle || 'Unknown',
                color: 'Warning',
                size: 'Medium',
              },
            ],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Request ID:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: approvalId || 'N/A',
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Justification:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: justification || 'No justification provided',
                wrap: true,
                size: 'Medium',
              },
            ],
          },
        ],
      },
      // Status banner
      {
        type: 'Container',
        style: status === 'approved' ? 'good' : 'attention',
        items: [
          {
            type: 'TextBlock',
            text: statusText,
            weight: 'Bolder',
            size: 'Large',
            horizontalAlignment: 'Center',
          },
          {
            type: 'TextBlock',
            text: `By: ${approverName || 'Unknown'} at ${time}`,
            size: 'Small',
            horizontalAlignment: 'Center',
            isSubtle: true,
          },
        ],
      },
    ],
    // No actions - request has been processed
    actions: [],
  };
  
  // Add approval details if approved
  if (status === 'approved') {
    const detailsItems = [];
    
    // Add record name first
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
      // Insert before the status banner
      card.body.splice(2, 0, {
        type: 'Container',
        spacing: 'Medium',
        items: detailsItems,
      });
    }
  }
  
  return card;
}

/**
 * Build a folder approval card with status (for updating existing card)
 * Shows the same card but with APPROVED/DENIED status and no action buttons
 */
function buildFolderApprovalCardWithStatus({
  approvalId,
  requesterName,
  requesterEmail,
  folderName,
  justification,
  status, // 'approved' or 'denied'
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
      // Header
      {
        type: 'TextBlock',
        text: 'Folder Access Request',
        weight: 'Bolder',
        size: 'ExtraLarge',
      },
      // Two-column layout for details
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Requester:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: requesterName || 'Unknown',
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Folder:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: folderName || 'Unknown',
                color: 'Warning',
                size: 'Medium',
              },
            ],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Request ID:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: approvalId || 'N/A',
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Justification:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: justification || 'No justification provided',
                wrap: true,
                size: 'Medium',
              },
            ],
          },
        ],
      },
      // Status banner
      {
        type: 'Container',
        style: status === 'approved' ? 'good' : 'attention',
        items: [
          {
            type: 'TextBlock',
            text: statusText,
            weight: 'Bolder',
            size: 'Large',
            horizontalAlignment: 'Center',
          },
          {
            type: 'TextBlock',
            text: `By: ${approverName || 'Unknown'} at ${time}`,
            size: 'Small',
            horizontalAlignment: 'Center',
            isSubtle: true,
          },
        ],
      },
    ],
    // No actions - request has been processed
    actions: [],
  };
  
  // Add approval details if approved
  if (status === 'approved') {
    const detailsItems = [];
    
    // Add folder name first
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
      // Insert before the status banner
      card.body.splice(2, 0, {
        type: 'Container',
        spacing: 'Medium',
        items: detailsItems,
      });
    }
  }
  
  return card;
}

/**
 * Format permission value to human-readable label
 */
function formatPermissionLabel(permission) {
  const labels = {
    'view_only': 'View Only',
    'can_edit': 'Can Edit',
    'can_share': 'Can Share',
    'edit_and_share': 'Edit & Share',
    'change_owner': 'Change Owner',
  };
  return labels[permission] || permission;
}

/**
 * Format folder permission value to human-readable label
 */
function formatFolderPermissionLabel(permission) {
  const labels = {
    'no_permissions': 'No User Permissions',
    'manage_users': 'Manage Users',
    'manage_records': 'Manage Records',
    'manage_all': 'Manage Records and Users',
  };
  return labels[permission] || permission;
}

/**
 * Build a record confirmation card (after search/selection in task module)
 * Shows selected record details with Approve/Deny buttons using Action.Execute
 * This allows the card to be updated via invoke response when admin clicks Approve
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
      // Header with "Ready to Approve" indicator
      {
        type: 'TextBlock',
        text: 'Record Selected - Ready to Approve',
        weight: 'Bolder',
        size: 'Large',
        color: 'Good',
      },
      // Two-column layout for details
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Requester:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: requesterName || 'Unknown',
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Selected Record:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: recordTitle || recordUid,
                color: 'Good',
                size: 'Medium',
                weight: 'Bolder',
              },
            ],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Request ID:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: approvalId || 'N/A',
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Justification:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: justification || 'No justification provided',
                wrap: true,
                size: 'Medium',
              },
            ],
          },
        ],
      },
      // Access details section
      {
        type: 'Container',
        separator: true,
        spacing: 'Medium',
        items: [
          {
            type: 'TextBlock',
            text: 'Access Configuration',
            weight: 'Bolder',
            size: 'Medium',
          },
          {
            type: 'FactSet',
            facts: [
              { title: 'Permission:', value: formatPermissionLabel(permission) },
              { title: 'Duration:', value: duration === 'permanent' ? 'Permanent' : duration },
              { title: 'For User:', value: requesterEmail },
            ],
          },
        ],
      },
      // Instructions
      {
        type: 'TextBlock',
        text: 'Click **Approve** to grant access or **Deny** to reject this request.',
        wrap: true,
        isSubtle: true,
        size: 'Small',
        spacing: 'Medium',
      },
    ],
    actions: [
      {
        type: 'Action.Execute',
        title: 'Approve',
        style: 'positive',
        verb: 'approve_record',
        data: {
          action: 'approve_record',
          approvalId: approvalId,
          recordUid: recordUid,
          recordTitle: recordTitle,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterName: requesterName,
          justification: justification,
          permission: permission,
          duration: duration,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny',
        style: 'destructive',
        verb: 'deny_record',
        data: {
          action: 'deny_record',
          approvalId: approvalId,
          recordUid: recordUid,
          recordTitle: recordTitle,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterName: requesterName,
          justification: justification,
        },
      },
    ],
  };
}

/**
 * Build a folder confirmation card (after search/selection in task module)
 * Shows selected folder details with Approve/Deny buttons using Action.Execute
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
      // Header with "Ready to Approve" indicator
      {
        type: 'TextBlock',
        text: 'Folder Selected - Ready to Approve',
        weight: 'Bolder',
        size: 'Large',
        color: 'Good',
      },
      // Two-column layout for details
      {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Requester:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: requesterName || 'Unknown',
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Selected Folder:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: folderName || folderUid,
                color: 'Good',
                size: 'Medium',
                weight: 'Bolder',
              },
            ],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              {
                type: 'TextBlock',
                text: 'Request ID:',
                weight: 'Bolder',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: approvalId || 'N/A',
                color: 'Warning',
                size: 'Medium',
              },
              {
                type: 'TextBlock',
                text: 'Justification:',
                weight: 'Bolder',
                size: 'Medium',
                spacing: 'Medium',
              },
              {
                type: 'TextBlock',
                text: justification || 'No justification provided',
                wrap: true,
                size: 'Medium',
              },
            ],
          },
        ],
      },
      // Access details section
      {
        type: 'Container',
        separator: true,
        spacing: 'Medium',
        items: [
          {
            type: 'TextBlock',
            text: 'Access Configuration',
            weight: 'Bolder',
            size: 'Medium',
          },
          {
            type: 'FactSet',
            facts: [
              { title: 'Permission:', value: formatFolderPermissionLabel(permission) },
              { title: 'Duration:', value: duration === 'permanent' ? 'Permanent' : duration },
              { title: 'For User:', value: requesterEmail },
            ],
          },
        ],
      },
      // Instructions
      {
        type: 'TextBlock',
        text: 'Click **Approve** to grant access or **Deny** to reject this request.',
        wrap: true,
        isSubtle: true,
        size: 'Small',
        spacing: 'Medium',
      },
    ],
    actions: [
      {
        type: 'Action.Execute',
        title: 'Approve',
        style: 'positive',
        verb: 'approve_folder',
        data: {
          action: 'approve_folder',
          approvalId: approvalId,
          folderUid: folderUid,
          folderName: folderName,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterName: requesterName,
          justification: justification,
          permission: permission,
          duration: duration,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny',
        style: 'destructive',
        verb: 'deny_folder',
        data: {
          action: 'deny_folder',
          approvalId: approvalId,
          folderUid: folderUid,
          folderName: folderName,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterName: requesterName,
          justification: justification,
        },
      },
    ],
  };
}

// Aliases for different naming conventions
const createRecordApprovalCard = buildRecordApprovalCard;
const createFolderApprovalCard = buildFolderApprovalCard;
const createShareApprovalCard = buildOneTimeShareApprovalCard;

module.exports = {
  buildRecordApprovalCard,
  buildRecordApprovalCardWithStatus,
  buildRecordConfirmationCard,
  buildRecordSearchResultsCard,
  buildFolderApprovalCard,
  buildFolderApprovalCardWithStatus,
  buildFolderConfirmationCard,
  buildFolderSearchResultsCard,
  buildShareSearchResultsCard,
  buildOneTimeShareApprovalCard,
  buildOneTimeShareApprovalCardWithStatus,
  createRecordApprovalCard,
  createFolderApprovalCard,
  createShareApprovalCard,
  RECORD_PERMISSIONS,
  FOLDER_PERMISSIONS,
  DURATION_OPTIONS,
  formatPermissionLabel,
  formatFolderPermissionLabel,
};
