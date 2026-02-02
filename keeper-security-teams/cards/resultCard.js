/**
 * Adaptive Card builders for approval results and notifications
 * 
 * These cards are sent to users to notify them of approval decisions,
 * share links, and other results.
 */

/**
 * Build an Adaptive Card for approval result notification
 */
function buildApprovalResultCard({
  approved,
  approverName,
  itemName,
  itemType = 'record',
  permission,
  duration,
  expiresAt,
  reason,
}) {
  const icon = approved ? '✅' : '❌';
  const title = approved ? 'Access Request Approved' : 'Access Request Denied';
  const color = approved ? 'Good' : 'Attention';
  
  const body = [
    {
      type: 'Container',
      style: approved ? 'good' : 'attention',
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
                  text: icon,
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
                  text: title,
                  weight: 'Bolder',
                  size: 'Large',
                  color: color,
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
            { title: itemType === 'folder' ? 'Folder' : 'Record', value: itemName },
            { title: 'Reviewed by', value: approverName },
          ],
        },
      ],
    },
  ];
  
  if (approved) {
    body.push({
      type: 'Container',
      items: [
        {
          type: 'FactSet',
          facts: [
            { title: 'Permission', value: formatPermission(permission) },
            { title: 'Duration', value: duration || 'Permanent' },
            ...(expiresAt ? [{ title: 'Expires', value: formatDate(expiresAt) }] : []),
          ],
        },
      ],
    });
  } else if (reason) {
    body.push({
      type: 'Container',
      items: [
        {
          type: 'TextBlock',
          text: 'Reason',
          weight: 'Bolder',
        },
        {
          type: 'TextBlock',
          text: reason,
          wrap: true,
          color: 'Attention',
        },
      ],
    });
  }
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: body,
  };
}

/**
 * Build an Adaptive Card for one-time share link result
 */
function buildShareResultCard({
  success = true,
  recordTitle,
  shareUrl,
  expiresAt,
  error,
}) {
  if (!success || error) {
    return buildErrorCard('Share Creation Failed', error || 'Unable to create share link');
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
                    text: 'One-Time Share Link Created',
                    weight: 'Bolder',
                    size: 'Large',
                    color: 'Good',
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
              { title: 'Record', value: recordTitle },
              { title: 'Expires', value: formatDate(expiresAt) },
            ],
          },
        ],
      },
      {
        type: 'Container',
        items: [
          {
            type: 'TextBlock',
            text: 'Share URL',
            weight: 'Bolder',
          },
          {
            type: 'TextBlock',
            text: shareUrl,
            wrap: true,
            color: 'Accent',
          },
        ],
      },
      {
        type: 'Container',
        items: [
          {
            type: 'TextBlock',
            text: '⚠️ This link can only be used once and will expire after first access or at the expiration time.',
            wrap: true,
            size: 'Small',
            isSubtle: true,
          },
        ],
      },
    ],
    actions: [
      {
        type: 'Action.OpenUrl',
        title: '🔗 Open Share Link',
        url: shareUrl,
      },
    ],
  };
}

/**
 * Build an Adaptive Card for search results
 */
function buildSearchResultsCard({
  searchType,
  query,
  results,
}) {
  const icon = searchType === 'folder' ? '📁' : '🔐';
  const title = searchType === 'folder' ? 'Folder Search Results' : 'Record Search Results';
  
  const resultItems = results.map(item => {
    if (searchType === 'folder') {
      return {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'auto',
            items: [{ type: 'TextBlock', text: '📁' }],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              { type: 'TextBlock', text: item.name, weight: 'Bolder' },
              { type: 'TextBlock', text: 'UID: ' + item.uid, size: 'Small', isSubtle: true },
            ],
          },
        ],
      };
    } else {
      return {
        type: 'ColumnSet',
        columns: [
          {
            type: 'Column',
            width: 'auto',
            items: [{ type: 'TextBlock', text: getRecordIcon(item.recordType) }],
          },
          {
            type: 'Column',
            width: 'stretch',
            items: [
              { type: 'TextBlock', text: item.title, weight: 'Bolder' },
              { type: 'TextBlock', text: 'UID: ' + item.uid + ' | Type: ' + item.recordType, size: 'Small', isSubtle: true },
            ],
          },
        ],
      };
    }
  });
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        items: [
          {
            type: 'ColumnSet',
            columns: [
              {
                type: 'Column',
                width: 'auto',
                items: [{ type: 'TextBlock', text: icon, size: 'Large' }],
              },
              {
                type: 'Column',
                width: 'stretch',
                items: [
                  { type: 'TextBlock', text: title, weight: 'Bolder', size: 'Medium' },
                  { type: 'TextBlock', text: 'Query: "' + query + '" | Found: ' + results.length, size: 'Small', isSubtle: true },
                ],
              },
            ],
          },
        ],
      },
      {
        type: 'Container',
        separator: true,
        items: resultItems,
      },
    ],
  };
}

/**
 * Build an Adaptive Card for help
 */
function buildHelpCard() {
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
                items: [{ type: 'TextBlock', text: '🔐', size: 'ExtraLarge' }],
              },
              {
                type: 'Column',
                width: 'stretch',
                items: [
                  { type: 'TextBlock', text: 'Keeper Security Bot', weight: 'Bolder', size: 'Large' },
                  { type: 'TextBlock', text: 'Secure credential access for Teams', size: 'Small', isSubtle: true },
                ],
              },
            ],
          },
        ],
      },
      {
        type: 'Container',
        items: [
          { type: 'TextBlock', text: '📋 Request Access', weight: 'Bolder', color: 'Accent' },
          { type: 'TextBlock', text: '• `keeper-request-record <name> <reason>` - Request record access', wrap: true },
          { type: 'TextBlock', text: '• `keeper-request-folder <name> <reason>` - Request folder access', wrap: true },
        ],
      },
      {
        type: 'Container',
        items: [
          { type: 'TextBlock', text: '🔗 Share', weight: 'Bolder', color: 'Accent' },
          { type: 'TextBlock', text: '• `keeper-one-time-share <name> [reason]` - Request one-time share link', wrap: true },
        ],
      },
      {
        type: 'Container',
        items: [
          { type: 'TextBlock', text: '🔍 Search', weight: 'Bolder', color: 'Accent' },
          { type: 'TextBlock', text: '• `search-records <query>` - Search records', wrap: true },
          { type: 'TextBlock', text: '• `search-folders <query>` - Search folders', wrap: true },
        ],
      },
      {
        type: 'Container',
        items: [
          { type: 'TextBlock', text: '⚙️ Other', weight: 'Bolder', color: 'Accent' },
          { type: 'TextBlock', text: '• `status` - Check connection', wrap: true },
          { type: 'TextBlock', text: '• `help` - Show this help', wrap: true },
        ],
      },
    ],
  };
}

/**
 * Build an Adaptive Card for error notification
 */
function buildErrorCard(title, message) {
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
                  { type: 'TextBlock', text: title, weight: 'Bolder', size: 'Medium', color: 'Attention' },
                ],
              },
            ],
          },
        ],
      },
      {
        type: 'Container',
        items: [
          { type: 'TextBlock', text: message, wrap: true },
        ],
      },
    ],
  };
}

/**
 * Build an Adaptive Card showing the approved message update
 */
function buildApprovedMessageCard(approverName, permission, duration, itemName, itemType) {
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
                  { type: 'TextBlock', text: 'Request Approved', weight: 'Bolder', size: 'Medium', color: 'Good' },
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
              { title: itemType === 'folder' ? 'Folder' : 'Record', value: itemName || 'N/A' },
              { title: 'Approved by', value: approverName },
              { title: 'Permission', value: formatPermission(permission) },
              { title: 'Duration', value: duration || 'Permanent' },
            ],
          },
        ],
      },
    ],
  };
}

/**
 * Build an Adaptive Card showing the denied message update
 */
function buildDeniedMessageCard(approverName, reason, itemName, itemType) {
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
                  { type: 'TextBlock', text: 'Request Denied', weight: 'Bolder', size: 'Medium', color: 'Attention' },
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
              { title: itemType === 'folder' ? 'Folder' : 'Record', value: itemName || 'N/A' },
              { title: 'Denied by', value: approverName },
              ...(reason ? [{ title: 'Reason', value: reason }] : []),
            ],
          },
        ],
      },
    ],
  };
}

// Helper functions
function formatPermission(permission) {
  const map = {
    'view_only': 'View Only',
    'can_edit': 'Can Edit',
    'can_share': 'Can Share',
    'edit_and_share': 'Edit & Share',
    'change_owner': 'Owner',
    'no_permissions': 'No Permissions',
    'manage_users': 'Manage Users',
    'manage_records': 'Manage Records',
    'manage_all': 'Manage All',
  };
  return map[permission] || permission;
}

function formatDate(dateStr) {
  try {
    const date = new Date(dateStr);
    return date.toLocaleString();
  } catch (e) {
    return dateStr;
  }
}

function getRecordIcon(recordType) {
  const icons = {
    'login': '🔑',
    'bankAccount': '🏦',
    'bankCard': '💳',
    'sshKeys': '🔐',
    'encryptedNotes': '📝',
    'file': '📄',
    'serverCredentials': '🖥️',
    'databaseCredentials': '🗄️',
  };
  return icons[recordType] || '🔐';
}

/**
 * Build a notification card for the requester after approval/denial
 * This is sent as a DM to the person who originally requested access
 */
function buildRequesterNotificationCard({
  approved,
  recordTitle,
  permission,
  duration,
  expiresAt,
  approverName,
  denialReason,
  itemType = 'record',
}) {
  const statusText = approved ? 'Your Access Request Has Been Approved!' : 'Your Access Request Has Been Denied';
  const containerStyle = approved ? 'good' : 'attention';
  const itemLabel = itemType === 'folder' ? 'Folder' : 'Record';
  
  const body = [
    // Header with status
    {
      type: 'Container',
      style: containerStyle,
      items: [
        {
          type: 'TextBlock',
          text: statusText,
          weight: 'Bolder',
          size: 'Large',
          wrap: true,
        },
      ],
    },
    // Record info
    {
      type: 'Container',
      spacing: 'Medium',
      items: [
        {
          type: 'ColumnSet',
          columns: [
            {
              type: 'Column',
              width: 'auto',
              items: [{ type: 'TextBlock', text: `${itemLabel}:`, weight: 'Bolder' }],
            },
            {
              type: 'Column',
              width: 'stretch',
              items: [{ type: 'TextBlock', text: recordTitle || 'Unknown', color: 'Accent', wrap: true }],
            },
          ],
        },
        {
          type: 'ColumnSet',
          columns: [
            {
              type: 'Column',
              width: 'auto',
              items: [{ type: 'TextBlock', text: 'Reviewed by:', weight: 'Bolder' }],
            },
            {
              type: 'Column',
              width: 'stretch',
              items: [{ type: 'TextBlock', text: approverName || 'Unknown' }],
            },
          ],
        },
      ],
    },
  ];
  
  if (approved) {
    // Add approval details
    body.push({
      type: 'Container',
      spacing: 'Medium',
      separator: true,
      items: [
        {
          type: 'TextBlock',
          text: 'Access Details',
          weight: 'Bolder',
          color: 'Good',
        },
        {
          type: 'FactSet',
          facts: [
            { title: 'Permission', value: formatPermission(permission) },
            { title: 'Duration', value: duration || 'Permanent' },
            ...(expiresAt ? [{ title: 'Expires', value: expiresAt }] : [{ title: 'Expires', value: 'Never (Permanent)' }]),
          ],
        },
      ],
    });
    
    // Add helpful message
    body.push({
      type: 'Container',
      spacing: 'Medium',
      items: [
        {
          type: 'TextBlock',
          text: `💡 You now have access to this ${itemLabel.toLowerCase()} in your Keeper vault.`,
          wrap: true,
          isSubtle: true,
          size: 'Small',
        },
      ],
    });
  } else {
    // Add denial info
    if (denialReason) {
      body.push({
        type: 'Container',
        spacing: 'Medium',
        separator: true,
        items: [
          {
            type: 'TextBlock',
            text: 'Reason:',
            weight: 'Bolder',
          },
          {
            type: 'TextBlock',
            text: denialReason,
            wrap: true,
            color: 'Attention',
          },
        ],
      });
    }
    
    // Add helpful message for denial
    body.push({
      type: 'Container',
      spacing: 'Medium',
      items: [
        {
          type: 'TextBlock',
          text: '💡 You can submit a new request with additional justification if needed.',
          wrap: true,
          isSubtle: true,
          size: 'Small',
        },
      ],
    });
  }
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.2',
    body: body,
  };
}

// Aliases
const createShareResultCard = buildShareResultCard;
const createSearchResultsCard = buildSearchResultsCard;
const createHelpCard = buildHelpCard;

module.exports = {
  buildApprovalResultCard,
  buildShareResultCard,
  buildSearchResultsCard,
  buildHelpCard,
  buildErrorCard,
  buildApprovedMessageCard,
  buildDeniedMessageCard,
  buildRequesterNotificationCard,
  createShareResultCard,
  createSearchResultsCard,
  createHelpCard,
  formatPermission,
  formatDate,
  getRecordIcon,
};
