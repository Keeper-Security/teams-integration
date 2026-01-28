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
        type: 'Action.Submit',
        title: 'Approve',
        style: 'positive',
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
        type: 'Action.Submit',
        title: 'Deny',
        style: 'destructive',
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
        type: 'Action.Submit',
        title: 'Deny Request',
        style: 'destructive',
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
 */
function buildFolderApprovalCard({
  approvalId,
  requesterName,
  requesterId,
  requesterEmail,
  folderName,
  folderUid,
  folderType = 'shared_folder',
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
                    text: '📁',
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
                    text: 'Folder Access Request',
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
              { title: 'Folder', value: folderName },
              { title: 'Folder UID', value: folderUid },
              { title: 'Type', value: folderType },
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
            text: 'Permission Level',
            weight: 'Bolder',
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
          },
          {
            type: 'Input.ChoiceSet',
            id: 'duration',
            value: '24h',
            choices: DURATION_OPTIONS,
          },
        ],
      },
    ],
    actions: [
      {
        type: 'Action.Submit',
        title: '✅ Approve',
        style: 'positive',
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
        type: 'Action.Submit',
        title: '❌ Deny',
        style: 'destructive',
        data: {
          action: 'deny_folder',
          approvalId: approvalId,
          folderUid: folderUid,
          folderName: folderName,
          requesterId: requesterId,
          requesterName: requesterName,
        },
      },
    ],
  };
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

// Aliases for different naming conventions
const createRecordApprovalCard = buildRecordApprovalCard;
const createFolderApprovalCard = buildFolderApprovalCard;
const createShareApprovalCard = buildOneTimeShareApprovalCard;

module.exports = {
  buildRecordApprovalCard,
  buildFolderApprovalCard,
  buildOneTimeShareApprovalCard,
  createRecordApprovalCard,
  createFolderApprovalCard,
  createShareApprovalCard,
  RECORD_PERMISSIONS,
  FOLDER_PERMISSIONS,
  DURATION_OPTIONS,
};
