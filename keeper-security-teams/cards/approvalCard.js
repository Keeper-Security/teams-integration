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
 * Slack-style layout with two columns
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
    // Description - show action required with search button inline
    card.body.push({
      type: 'ColumnSet',
      spacing: 'Large',
      columns: [
        {
          type: 'Column',
          width: 'stretch',
          verticalContentAlignment: 'Center',
          items: [
            {
              type: 'TextBlock',
              text: '**Action Required:** Approver must search for the correct record',
              wrap: true,
              size: 'Medium',
            },
          ],
        },
        {
          type: 'Column',
          width: 'auto',
          items: [
            {
              type: 'ActionSet',
              actions: [
                {
                  type: 'Action.Submit',
                  title: '🔍 Search Records',
                  data: {
                    action: 'search_records',
                    approvalId: approvalId,
                    identifier: identifier || recordTitle,
                    recordTitle: recordTitle,
                    requesterId: requesterId,
                    requesterEmail: requesterEmail,
                    requesterAadObjectId: requesterAadObjectId,
                    requesterName: requesterName,
                    justification: justification,
                    msteams: {
                      type: 'task/fetch',
                    },
                  },
                },
              ],
            },
          ],
        },
      ],
    });
    
    // Only Deny button at bottom
    card.actions = [
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
          requesterName: requesterName,
        },
      },
    ];
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
    // Description - show action required with search button inline
    card.body.push({
      type: 'ColumnSet',
      spacing: 'Large',
      columns: [
        {
          type: 'Column',
          width: 'stretch',
          verticalContentAlignment: 'Center',
          items: [
            {
              type: 'TextBlock',
              text: '**Action Required:** Approver must search for the correct folder',
              wrap: true,
              size: 'Medium',
            },
          ],
        },
        {
          type: 'Column',
          width: 'auto',
          items: [
            {
              type: 'ActionSet',
              actions: [
                {
                  type: 'Action.Submit',
                  title: '🔍 Search Folders',
                  data: {
                    action: 'search_folders',
                    approvalId: approvalId,
                    identifier: identifier || folderName,
                    folderName: folderName,
                    requesterId: requesterId,
                    requesterEmail: requesterEmail,
                    requesterAadObjectId: requesterAadObjectId,
                    requesterName: requesterName,
                    justification: justification,
                    msteams: {
                      type: 'task/fetch',
                    },
                  },
                },
              ],
            },
          ],
        },
      ],
    });
    
    // Only Deny button at bottom
    card.actions = [
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
          requesterName: requesterName,
        },
      },
    ];
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
  recordTitle,
  recordUid,
  justification,
}) {
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: 'emphasis',
        items: [
          {
            type: 'ColumnSet',
            columns: [
              {
                type: 'Column',
                width: 'auto',
                items: [
                  {
                    type: 'TextBlock',
                    text: '🔗',
                    size: 'ExtraLarge',
                  },
                ],
              },
              {
                type: 'Column',
                width: 'stretch',
                items: [
                  {
                    type: 'TextBlock',
                    text: 'One-Time Share Request',
                    weight: 'Bolder',
                    size: 'Large',
                    color: 'Accent',
                  },
                  {
                    type: 'TextBlock',
                    text: 'ID: ' + approvalId,
                    size: 'Small',
                    isSubtle: true,
                  },
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
              { title: 'Requester', value: requesterName },
              { title: 'Record', value: recordTitle },
              { title: 'Record UID', value: recordUid },
            ],
          },
        ],
      },
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
            color: justification ? 'Default' : 'Attention',
          },
        ],
      },
      {
        type: 'Container',
        separator: true,
        items: [
          {
            type: 'TextBlock',
            text: 'Share Duration',
            weight: 'Bolder',
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
            title: 'Allow editing',
            value: 'false',
          },
        ],
      },
    ],
    actions: [
      {
        type: 'Action.Submit',
        title: '✅ Approve & Create Link',
        style: 'positive',
        data: {
          action: 'approve_share',
          approvalId: approvalId,
          recordUid: recordUid,
          recordTitle: recordTitle,
          requesterId: requesterId,
          requesterEmail: requesterEmail,
          requesterName: requesterName,
        },
      },
      {
        type: 'Action.Submit',
        title: '❌ Deny',
        style: 'destructive',
        data: {
          action: 'deny_share',
          approvalId: approvalId,
          recordUid: recordUid,
          recordTitle: recordTitle,
          requesterId: requesterId,
          requesterName: requesterName,
        },
      },
    ],
  };
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

// Aliases for different naming conventions
const createRecordApprovalCard = buildRecordApprovalCard;
const createFolderApprovalCard = buildFolderApprovalCard;
const createShareApprovalCard = buildOneTimeShareApprovalCard;

module.exports = {
  buildRecordApprovalCard,
  buildRecordApprovalCardWithStatus,
  buildFolderApprovalCard,
  buildFolderApprovalCardWithStatus,
  buildOneTimeShareApprovalCard,
  createRecordApprovalCard,
  createFolderApprovalCard,
  createShareApprovalCard,
  RECORD_PERMISSIONS,
  FOLDER_PERMISSIONS,
  DURATION_OPTIONS,
  formatPermissionLabel,
  formatFolderPermissionLabel,
};
