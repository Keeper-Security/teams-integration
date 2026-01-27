/**
 * Task Module Handler
 * Handles Teams Task Modules for record search and selection
 * 
 * Similar to Slack's modal flow:
 * 1. Approver clicks "Search Records" button
 * 2. Task module opens with loading state
 * 3. Search executes (async)
 * 4. Modal updated with results (radio buttons, permission/duration selectors)
 * 5. Approver selects record and approves
 */

const keeperClient = require('../services/keeperClient');

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
  { title: 'Can Share (Permanent)', value: 'can_share' },
  { title: 'Edit & Share (Permanent)', value: 'edit_and_share' },
  { title: 'Change Owner (Permanent)', value: 'change_owner' },
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
 * Handle task/fetch - Show search modal
 * @param {Object} context - Teams turn context
 * @param {Object} activity - Activity object (from invoke or task/fetch event)
 * @returns {Promise<Object>} - Task module response
 */
async function handleTaskFetch(context, activity) {
  // activity might be the full activity object or just the value
  let requestData = {};
  
  console.log('[TaskModule] Raw activity structure:', {
    hasValue: !!activity.value,
    hasData: !!activity.data,
    valueKeys: activity.value ? Object.keys(activity.value) : [],
    activityKeys: Object.keys(activity)
  });
  
  // Try to extract data from different possible structures
  if (activity.value) {
    // If it's an invoke activity, the data is in activity.value
    // When Teams converts cardAction with msteams to invoke, the structure is:
    // activity.value = { data: { ...card action data... } }
    if (activity.value.data) {
      requestData = activity.value.data;
    } else if (activity.value.value && activity.value.value.data) {
      requestData = activity.value.value.data;
    } else {
      // Fallback: use activity.value directly (might be the card action data)
      requestData = activity.value;
    }
  } else if (activity.data) {
    // If data is directly on activity
    requestData = activity.data;
  } else {
    // Fallback: use activity itself
    requestData = activity;
  }
  
  // Extract fields - handle both direct fields and nested structure
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
    
    // Build approval context from card action data
    approvalContext = {
      approvalId: requestData.approvalId,
      requesterEmail: requestData.requesterEmail,
      requesterName: requestData.requesterName,
      requesterId: requestData.requesterId,
      requesterAadObjectId: requestData.requesterAadObjectId, // Store AAD Object ID for email fetching fallback
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
    // Show loading state first
    return buildSearchModal(searchQuery, searchApprovalId, true, contextData);
  }

  return buildSearchModal(searchQuery, searchApprovalId, false, contextData);
}

/**
 * Handle task/submit - Process search actions
 * @param {Object} context - Teams turn context
 * @param {Object} activity - Activity object (from invoke or task/submit event)
 * @returns {Promise<Object>} - Task module response
 */
async function handleTaskSubmit(context, activity) {
  // Extract data from activity
  let submitData = {};
  
  if (activity.value) {
    // If it's an invoke activity, the data is in activity.value
    if (activity.value.data) {
      submitData = activity.value.data;
    } else if (activity.value.value && activity.value.value.data) {
      submitData = activity.value.value.data;
    } else {
      submitData = activity.value;
    }
  } else if (activity.data) {
    submitData = activity.data;
  } else {
    submitData = activity;
  }
  
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
  
  const { action, searchQuery, selectedUid, approvalId, permission, duration } = submitData;

  console.log('[TaskModule] Submit:', { 
    action, 
    searchQuery, 
    selectedUid, 
    approvalId,
    permission,
    duration,
    hasApprovalContext: !!approvalContext
  });

  if (action === 'search' || action === 'refine_search') {
    // User clicked search/refine - return updated results
    return await handleSearchAction(searchQuery, approvalId, approvalContext || {});
  } else if (action === 'select_and_approve') {
    // User selected record and clicked approve
    return await handleSelectAndApprove(context, selectedUid, approvalId, permission, duration, approvalContext || {});
  }

  return {
    task: {
      type: 'message',
      value: 'Invalid action: ' + (action || 'none'),
    },
  };
}

/**
 * Build search modal (loading or with results)
 */
function buildSearchModal(query, approvalId, isLoading = false, approvalContext = {}) {
  const card = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'TextBlock',
        text: '🔍 Search Records',
        weight: 'Bolder',
        size: 'Large',
      },
      {
        type: 'TextBlock',
        text: 'Search for the correct record to approve access',
        wrap: true,
        isSubtle: true,
      },
      {
        type: 'Input.Text',
        id: 'searchQuery',
        label: 'Search Query',
        placeholder: 'Enter UID or record name...',
        value: query || '',
        isRequired: true,
      },
      {
        type: 'Input.ChoiceSet',
        id: 'permission',
        label: 'Permission Level',
        choices: RECORD_PERMISSIONS,
        value: 'view_only',
      },
      {
        type: 'Input.ChoiceSet',
        id: 'duration',
        label: 'Duration',
        choices: DURATION_OPTIONS,
        value: '24h',
      },
    ],
    actions: [],
  };

  if (isLoading) {
    card.body.push({
      type: 'TextBlock',
      text: '🔄 Searching...',
      color: 'Accent',
      wrap: true,
    });
  }

  // Store approval context in hidden field
  card.body.push({
    type: 'Input.Text',
    id: 'approvalContext',
    isVisible: false,
    value: JSON.stringify(approvalContext),
  });

  card.actions.push({
    type: 'Action.Submit',
    title: isLoading ? 'Search' : '🔍 Refine Search',
    data: {
      action: isLoading ? 'search' : 'refine_search',
      approvalId: approvalId,
      // Include approvalContext in the submit data so it's preserved
      approvalContext: typeof approvalContext === 'object' ? JSON.stringify(approvalContext) : approvalContext,
    },
  });

  // Ensure the response format matches Teams expectations
  // Teams requires the card to be wrapped with contentType
  const taskModuleResponse = {
    task: {
      type: 'continue',
      value: {
        title: '🔍 Search Records',
        height: 500, // Use numeric value for better compatibility
        width: 600,  // Use numeric value for better compatibility
        card: {
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: card, // The Adaptive Card JSON
        },
      },
    },
  };
  
  console.log('[TaskModule] Returning task module response:', JSON.stringify(taskModuleResponse, null, 2).substring(0, 500));
  
  return taskModuleResponse;
}

/**
 * Handle search action - execute search and return results
 */
async function handleSearchAction(query, approvalId, approvalContext) {
  if (!query || !query.trim()) {
    return {
      task: {
        type: 'continue',
        value: {
          title: 'Search Error',
          card: {
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: {
              type: 'AdaptiveCard',
              version: '1.5',
              body: [
                {
                  type: 'TextBlock',
                  text: 'Please enter a search query',
                  color: 'Attention',
                },
              ],
            },
          },
        },
      },
    };
  }

  console.log('[TaskModule] Executing search for:', query);

  // Try UID first
  let record = await keeperClient.getRecordByUid(query.trim());
  let results = [];
  
  if (record) {
    results = [record];
    console.log('[TaskModule] Found record by UID:', record.uid);
  } else {
    // Search by name/description
    results = await keeperClient.searchRecords(query, 20);
    console.log('[TaskModule] Found', results.length, 'records by search');
  }

  // Ensure approvalContext is an object
  const contextData = approvalContext || {};
  
  // Build results section
  const card = buildSearchModal(query, approvalId, false, contextData);
  const resultItems = [];

  if (results.length === 0) {
    resultItems.push({
      type: 'TextBlock',
      text: `No records found matching "${query}"`,
      color: 'Attention',
      wrap: true,
    });
  } else {
    resultItems.push({
      type: 'TextBlock',
      text: `Found ${results.length} result(s):`,
      weight: 'Bolder',
      wrap: true,
    });

    // Add radio buttons for selection
    const choices = results.map((item) => ({
      title: `${item.title || item.uid} (${item.uid})`,
      value: item.uid,
    }));

    resultItems.push({
      type: 'Input.ChoiceSet',
      id: 'selectedUid',
      label: 'Select Record',
      style: 'expanded',
      choices: choices,
      isRequired: true,
    });

    resultItems.push({
      type: 'ActionSet',
      actions: [
        {
          type: 'Action.Submit',
          title: '✅ Approve Selected Record',
          style: 'positive',
          data: {
            action: 'select_and_approve',
            approvalId: approvalId,
            // Include approvalContext so it's available when approving
            approvalContext: typeof approvalContext === 'object' ? JSON.stringify(approvalContext) : approvalContext,
          },
        },
      ],
    });
  }

  // Insert results before actions (after duration selector)
  // Card structure is now: task.value.card.content (Adaptive Card)
  const adaptiveCard = card.task.value.card.content;
  const bodyIndex = adaptiveCard.body.length - 1; // Before hidden approvalContext field
  adaptiveCard.body.splice(bodyIndex, 0, {
    type: 'Container',
    id: 'searchResults',
    items: resultItems,
    separator: true,
  });

  return card;
}

/**
 * Handle select and approve - grant access and close modal
 */
async function handleSelectAndApprove(context, selectedUid, approvalId, permission, duration, approvalContext) {
  console.log('[TaskModule] Approving record:', { selectedUid, approvalId, permission, duration });

  // Get record details
  const record = await keeperClient.getRecordByUid(selectedUid);
  if (!record) {
    return {
      task: {
        type: 'message',
        value: `❌ Record not found: ${selectedUid}`,
      },
    };
  }

  // Parse duration
  const durationSeconds = DURATION_MAP[duration] ?? 86400;

  // Get requester email from approval context
  let requesterEmail = approvalContext.requesterEmail;
  
  // If email is missing, try to fetch it using stored AAD Object ID
  if (!requesterEmail && approvalContext.requesterAadObjectId) {
    console.log('[TaskModule] Email missing, fetching from Graph API using AAD Object ID...');
    try {
      const graphService = require('../services/graphService');
      requesterEmail = await graphService.getUserEmail(approvalContext.requesterAadObjectId);
      if (requesterEmail) {
        console.log(`[TaskModule] Successfully fetched email: ${requesterEmail}`);
      } else {
        console.warn('[TaskModule] Graph API returned no email');
      }
    } catch (error) {
      console.error('[TaskModule] Error fetching email from Graph API:', error.message);
    }
  }
  
  // If email is in old pseudo-format (@users.teams.ms), fetch real email
  if (requesterEmail && requesterEmail.endsWith('@users.teams.ms')) {
    console.log('[TaskModule] Detected old email format, fetching real email...');
    try {
      const aadObjectId = requesterEmail.replace('@users.teams.ms', '');
      const graphService = require('../services/graphService');
      const realEmail = await graphService.getUserEmail(aadObjectId);
      if (realEmail) {
        console.log(`[TaskModule] Fetched real email: ${realEmail} (was: ${requesterEmail})`);
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
        value: '❌ Error: Missing requester email. Cannot grant access.\n\n' +
               'The bot needs Microsoft Graph API permissions (User.Read) to fetch user emails.',
      },
    };
  }

  // Grant access
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
    // Update the approval card in the channel
    // We'll need to send a message to update the original card
    // For now, return success message
    
    const { sendAccessGrantedNotification } = require('../utils/helpers');
    const requesterId = approvalContext.requesterId;
    
    // Send DM to requester
    if (requesterId) {
      // Note: We'll need to implement proactive DM sending
      console.log('[TaskModule] Should send DM to requester:', requesterId);
    }

    return {
      task: {
        type: 'continue',
        value: {
          title: 'Request Approved',
          card: {
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: {
              type: 'AdaptiveCard',
              version: '1.5',
              body: [
                {
                  type: 'TextBlock',
                  text: '✅ Access Granted',
                  weight: 'Bolder',
                  size: 'Large',
                  color: 'Good',
                },
                {
                  type: 'TextBlock',
                  text: `Access has been granted for record: **${record.title || selectedUid}**`,
                  wrap: true,
                },
                {
                  type: 'TextBlock',
                  text: `• **Permission:** ${permission}`,
                  wrap: true,
                },
                {
                  type: 'TextBlock',
                  text: `• **Duration:** ${duration}`,
                  wrap: true,
                },
                {
                  type: 'TextBlock',
                  text: `• **Expires:** ${result.expiresAt || 'Permanent'}`,
                  wrap: true,
                  isSubtle: true,
                },
              ],
            },
          },
        },
      },
    };
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
