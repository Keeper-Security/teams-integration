/**
 * Task Module Handler
 * Handles Teams Task Modules for record search and selection
 * 
 * Similar to Slack's modal flow:
 * 1. Approver clicks "Search Records" button
 * 2. Task module opens with search results (auto-search)
 * 3. Approver selects record, permission, duration
 * 4. Approver clicks "Approve Access"
 * 
 * IMPORTANT: Use Adaptive Card version 1.2 for Task Module compatibility
 * Buttons must be in the card's main actions array (not nested in ColumnSet)
 */

const keeperClient = require('../services/keeperClient');
const cards = require('../cards');
const { getChannelService } = require('../services');

/**
 * Duration string to seconds mapping
 */
const DURATION_MAP = {
  '1h': 3600,
  '4h': 14400,
  '8h': 28800,
  '24h': 86400,
  '7d': 604800,
  '30d': 2592000,
  'permanent': null,
};

/**
 * Permission options for records
 */
const RECORD_PERMISSIONS = [
  { title: 'View Only', value: 'view_only' },
  { title: 'Can Edit', value: 'can_edit' },
  { title: 'Can Share', value: 'can_share' },
  { title: 'Edit & Share', value: 'edit_and_share' },
  { title: 'Change Owner', value: 'change_owner' },
];

/**
 * Duration options
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
 * Permissions that are permanent-only (no duration support)
 * Like Slack, these permissions don't support time limits
 */
const PERMANENT_ONLY_PERMISSIONS = ['can_share', 'edit_and_share', 'change_owner'];

/**
 * Handle task/fetch - Show search modal
 */
async function handleTaskFetch(context, activity) {
  let requestData = {};
  
  console.log('[TaskModule] Raw activity structure:', {
    hasValue: !!activity.value,
    hasData: !!activity.data,
    valueKeys: activity.value ? Object.keys(activity.value) : [],
    activityKeys: Object.keys(activity)
  });
  
  // Try to extract data from different possible structures
  if (activity.value) {
    if (activity.value.data) {
      requestData = activity.value.data;
    } else if (activity.value.value && activity.value.value.data) {
      requestData = activity.value.value.data;
    } else {
      requestData = activity.value;
    }
  } else if (activity.data) {
    requestData = activity.data;
  } else {
    requestData = activity;
  }
  
  // Extract fields
  let type = requestData.type;
  let query = requestData.query;
  let approvalId = requestData.approvalId;
  let identifier = requestData.identifier;
  let approvalContext = requestData.approvalContext;
  
  // If action is 'search_records', extract the data from the card action
  if (requestData.action === 'search_records') {
    type = 'search-record';
    query = requestData.identifier;
    approvalId = requestData.approvalId;
    identifier = requestData.identifier;
    
    approvalContext = {
      approvalId: requestData.approvalId,
      requesterEmail: requestData.requesterEmail,
      requesterName: requestData.requesterName,
      requesterId: requestData.requesterId,
      requesterAadObjectId: requestData.requesterAadObjectId,
      justification: requestData.justification,
      identifier: requestData.identifier,
      recordTitle: requestData.recordTitle,
    };
  }
  
  console.log('[TaskModule] Fetch:', { 
    type, 
    query, 
    approvalId, 
    identifier,
    hasApprovalContext: !!approvalContext,
    rawRequest: JSON.stringify(requestData).substring(0, 400)
  });

  const searchType = type || 'search-record';
  const searchQuery = query || identifier;
  const searchApprovalId = approvalId;
  const contextData = approvalContext || requestData;

  if (searchType === 'search-record' || searchType === 'search_records') {
    if (searchQuery && searchQuery.trim()) {
      console.log('[TaskModule] Auto-executing search on modal open for:', searchQuery);
      return await handleSearchAction(searchQuery, searchApprovalId, contextData, true);
    } else {
      return buildSearchModal('', searchApprovalId, false, contextData, [], true);
    }
  }

  return buildSearchModal(searchQuery, searchApprovalId, false, contextData, [], true);
}

/**
 * Handle task/submit - Process search actions
 */
async function handleTaskSubmit(context, activity) {
  // Simple data extraction - activity.value contains form data directly
  const submitData = activity.value?.data || activity.value || {};
  
  console.log('[TaskModule] Submit data:', JSON.stringify(submitData, null, 2));
  
  // Parse approvalContext if it's a JSON string
  let approvalContext = submitData.approvalContext;
  if (typeof approvalContext === 'string') {
    try {
      approvalContext = JSON.parse(approvalContext);
    } catch (e) {
      console.warn('[TaskModule] Failed to parse approvalContext:', e);
      approvalContext = {};
    }
  }
  
  // Parse cachedResults if it's a JSON string
  let cachedResults = submitData.cachedResults;
  if (typeof cachedResults === 'string') {
    try {
      cachedResults = JSON.parse(cachedResults);
    } catch (e) {
      console.warn('[TaskModule] Failed to parse cachedResults:', e);
      cachedResults = [];
    }
  }
  
  const { action, searchQuery, selectedUid, approvalId, permission, duration } = submitData;

  console.log('[TaskModule] Submit:', { 
    action, 
    searchQuery, 
    selectedUid, 
    approvalId,
    permission,
    duration,
    hasApprovalContext: !!approvalContext,
    hasCachedResults: !!cachedResults
  });

  if (action === 'search' || action === 'refine_search') {
    // Determine showDuration based on current permission
    const showDuration = !PERMANENT_ONLY_PERMISSIONS.includes(permission || 'view_only');
    return await handleSearchAction(searchQuery, approvalId, approvalContext || {}, showDuration, permission);
  } else if (action === 'select_and_approve') {
    return await handleSelectAndApprove(context, selectedUid, approvalId, permission, duration, approvalContext || {});
  } else if (action === 'close') {
    // Just close the task module
    return null;
  }

  return {
    task: {
      type: 'message',
      value: 'Invalid action: ' + (action || 'none'),
    },
  };
}

/**
 * Build search modal - Slack-style UI layout
 * IMPORTANT: Keep structure flat and simple for button compatibility
 */
function buildSearchModal(query, approvalId, isLoading = false, approvalContext = {}, results = [], showDuration = true, currentPermission = 'view_only', selectedUid = null) {
  
  const body = [];
  
  // Search Term section
  body.push({
    type: 'TextBlock',
    text: 'Search Term',
    weight: 'Bolder',
  });
  
  body.push({
    type: 'Input.Text',
    id: 'searchQuery',
    placeholder: 'Enter record name or UID...',
    value: query || '',
  });
  
  body.push({
    type: 'TextBlock',
    text: 'Modify the search term and click the Refine button below.',
    wrap: true,
    isSubtle: true,
    size: 'Small',
  });
  
  // Inline buttons with spacing: Refine Search + Create New Record
  body.push({
    type: 'ColumnSet',
    columns: [
      {
        type: 'Column',
        width: 'auto',
        items: [
          {
            type: 'ActionSet',
            actions: [
              {
                type: 'Action.Submit',
                title: 'Refine Search',
                data: { action: 'refine_search' },
              },
            ],
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
                title: 'Create New Record',
                style: 'positive',
                data: { action: 'create_new_record' },
              },
            ],
          },
        ],
      },
    ],
  });
  
  body.push({
    type: 'TextBlock',
    text: 'Or create a new record and share it',
    wrap: true,
    isSubtle: true,
    size: 'Small',
  });
  
  // Results section
  if (isLoading) {
    body.push({
      type: 'TextBlock',
      text: '🔄 Searching...',
      color: 'Accent',
      wrap: true,
    });
  } else if (results.length > 0) {
    body.push({
      type: 'TextBlock',
      text: `Showing ${results.length} result(s) for: **${query}**`,
      wrap: true,
    });
    
    // Record selection with required indicator
    body.push({
      type: 'TextBlock',
      text: 'Select record: *',
      weight: 'Bolder',
    });
    
    body.push({
      type: 'Input.ChoiceSet',
      id: 'selectedUid',
      style: 'expanded',
      choices: results.map(r => ({ 
        title: `${r.title} (${r.uid})`, 
        value: r.uid 
      })),
      value: selectedUid || '',
    });
    
    // Permission dropdown
    body.push({
      type: 'TextBlock',
      text: 'Select Permission Level',
      weight: 'Bolder',
    });
    
    body.push({
      type: 'Input.ChoiceSet',
      id: 'permission',
      choices: RECORD_PERMISSIONS,
      value: currentPermission,
    });
    
    // Duration section (conditional)
    if (showDuration) {
      body.push({
        type: 'TextBlock',
        text: 'Grant Access For',
        weight: 'Bolder',
      });
      
      body.push({
        type: 'Input.ChoiceSet',
        id: 'duration',
        choices: DURATION_OPTIONS,
        value: '1h',
      });
      
      body.push({
        type: 'TextBlock',
        text: 'Select how long the access should remain active.',
        wrap: true,
        isSubtle: true,
        size: 'Small',
      });
    } else {
      // Permanent access notice
      body.push({
        type: 'TextBlock',
        text: 'ℹ️ **Permanent Access** - This permission does not support time limits.',
        wrap: true,
        color: 'Accent',
      });
      
      // Hidden duration value
      body.push({
        type: 'Input.Text',
        id: 'duration',
        isVisible: false,
        value: 'permanent',
      });
    }
  } else if (query) {
    body.push({
      type: 'TextBlock',
      text: `❌ No records found matching "${query}"`,
      color: 'Attention',
      wrap: true,
    });
  }
  
  // Hidden context
  body.push({
    type: 'Input.Text',
    id: 'approvalContext',
    isVisible: false,
    value: JSON.stringify(approvalContext),
  });
  
  // Bottom action - only Approve Access (when results exist)
  const actions = [];
  if (results && results.length > 0) {
    actions.push({
      type: 'Action.Submit',
      title: 'Approve Access',
      style: 'positive',
      data: { action: 'select_and_approve' },
    });
  }
  
  const card = {
    type: 'AdaptiveCard',
    version: '1.2',
    body: body,
    actions: actions,
  };

  return {
    task: {
      type: 'continue',
      value: {
        title: 'Search Records',
        height: 600,
        width: 450,
        card: {
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: card,
        },
      },
    },
  };
}

/**
 * Handle search action - execute search and return results
 */
async function handleSearchAction(query, approvalId, approvalContext, showDuration = true, currentPermission = 'view_only') {
  if (!query || !query.trim()) {
    return buildSearchModal('', approvalId, false, approvalContext, [], showDuration, currentPermission);
  }

  console.log('[TaskModule] Executing search for:', query);

  let record = await keeperClient.getRecordByUid(query.trim());
  let results = [];
  
  if (record) {
    results = [{ uid: record.uid, title: record.title || record.uid }];
    console.log('[TaskModule] Found record by UID:', record.uid);
  } else {
    const searchResults = await keeperClient.searchRecords(query, 10);
    results = searchResults.map(r => ({ uid: r.uid, title: r.title || r.uid }));
    console.log('[TaskModule] Found', results.length, 'records by search');
  }

  return buildSearchModal(query, approvalId, false, approvalContext, results, showDuration, currentPermission);
}

/**
 * Handle select and approve - grant access and close modal
 */
async function handleSelectAndApprove(context, selectedUid, approvalId, permission, duration, approvalContext) {
  console.log('[TaskModule] Approving record:', { selectedUid, approvalId, permission, duration });

  if (!selectedUid) {
    return {
      task: {
        type: 'message',
        value: '❌ Please select a record first.',
      },
    };
  }

  // Force permanent duration for permanent-only permissions
  if (PERMANENT_ONLY_PERMISSIONS.includes(permission)) {
    console.log(`[TaskModule] Permission "${permission}" is permanent-only, forcing duration to permanent`);
    duration = 'permanent';
  }

  const record = await keeperClient.getRecordByUid(selectedUid);
  if (!record) {
    return {
      task: {
        type: 'message',
        value: `❌ Record not found: ${selectedUid}`,
      },
    };
  }

  const durationSeconds = DURATION_MAP[duration] ?? 86400;

  // Get requester email
  let requesterEmail = approvalContext.requesterEmail;
  
  if (!requesterEmail && approvalContext.requesterAadObjectId) {
    console.log('[TaskModule] Email missing, fetching from Graph API...');
    try {
      const graphService = require('../services/graphService');
      requesterEmail = await graphService.getUserEmail(approvalContext.requesterAadObjectId);
      if (requesterEmail) {
        console.log(`[TaskModule] Successfully fetched email: ${requesterEmail}`);
      }
    } catch (error) {
      console.error('[TaskModule] Error fetching email:', error.message);
    }
  }
  
  if (requesterEmail && requesterEmail.endsWith('@users.teams.ms')) {
    console.log('[TaskModule] Detected old email format, fetching real email...');
    try {
      const aadObjectId = requesterEmail.replace('@users.teams.ms', '');
      const graphService = require('../services/graphService');
      const realEmail = await graphService.getUserEmail(aadObjectId);
      if (realEmail) {
        console.log(`[TaskModule] Fetched real email: ${realEmail}`);
        requesterEmail = realEmail;
      }
    } catch (error) {
      console.error('[TaskModule] Error fetching real email:', error.message);
    }
  }
  
  if (!requesterEmail) {
    return {
      task: {
        type: 'message',
        value: '❌ Error: Missing requester email. Cannot grant access.\n\nThe bot needs Microsoft Graph API permissions (User.Read.All) to fetch user emails.',
      },
    };
  }

  console.log('[TaskModule] Granting access:', {
    recordUid: selectedUid,
    requesterEmail,
    permission,
    durationSeconds,
  });

  const result = await keeperClient.grantRecordAccess(
    selectedUid,
    requesterEmail,
    permission,
    durationSeconds
  );

  if (result.success) {
    const durationDisplay = duration === 'permanent' ? 'Permanent' : duration;
    
    // Format expiry date
    let expiresAtFormatted = null;
    if (result.expiresAt && duration !== 'permanent') {
      const expiryDate = new Date(result.expiresAt);
      expiresAtFormatted = expiryDate.toLocaleString('en-US', {
        month: '2-digit',
        day: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false,
      });
    }
    
    // Get approver info
    const approverName = context.activity.from?.name || 'Unknown';
    
    // Build the updated approval card with APPROVED status
    const updatedCard = cards.buildRecordApprovalCardWithStatus({
      approvalId: approvalContext.approvalId,
      requesterName: approvalContext.requesterName,
      requesterEmail: requesterEmail,
      recordTitle: record.title || selectedUid,
      justification: approvalContext.justification,
      status: 'approved',
      approverName: approverName,
      permission: permission,
      duration: durationDisplay,
      expiresAt: expiresAtFormatted,
    });
    
    // Send the updated card to the conversation
    try {
      await context.send({
        type: 'message',
        attachments: [{
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: updatedCard,
        }],
      });
      console.log('[TaskModule] Sent approval status card to conversation');
    } catch (sendError) {
      console.error('[TaskModule] Error sending approval status:', sendError.message);
    }
    
    // Send DM notification to the requester
    const requesterId = approvalContext.requesterId;
    if (requesterId) {
      try {
        const channelService = getChannelService();
        if (channelService) {
          const notificationCard = cards.buildRequesterNotificationCard({
            approved: true,
            recordTitle: record.title || selectedUid,
            permission: permission,
            duration: durationDisplay,
            expiresAt: expiresAtFormatted,
            approverName: approverName,
          });
          
          const notificationSent = await channelService.sendDirectMessage(requesterId, {
            type: 'message',
            attachments: [{
              contentType: 'application/vnd.microsoft.card.adaptive',
              content: notificationCard,
            }],
          });
          
          if (notificationSent) {
            console.log(`[TaskModule] Sent approval notification to requester: ${requesterId}`);
          } else {
            console.log(`[TaskModule] Could not send notification to requester (no reference stored)`);
          }
        }
      } catch (notifyError) {
        console.error('[TaskModule] Error sending requester notification:', notifyError.message);
      }
    }
    
    // Close the task module
    return null;
  } else {
    return {
      task: {
        type: 'message',
        value: `❌ Failed to grant access: ${result.error || 'Unknown error'}`,
      },
    };
  }
}

module.exports = {
  handleTaskFetch,
  handleTaskSubmit,
};
