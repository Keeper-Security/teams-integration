/**
 * Task Module Handler
 * Handles Teams Task Modules for record search and selection
 */

const keeperClient = require('../services/keeperClient');
const cards = require('../cards');
const { getChannelService, getApprovalActivityId } = require('../services');

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
 * Permission options for folders
 */
const FOLDER_PERMISSIONS = [
  { title: 'No User Permissions', value: 'no_permissions' },
  { title: 'Manage Users (Permanent)', value: 'manage_users' },
  { title: 'Manage Records', value: 'manage_records' },
  { title: 'Manage All (Permanent)', value: 'manage_all' },
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
 * Record permissions that are permanent-only (no duration support)
 * Like Slack, these permissions don't support time limits
 */
const PERMANENT_ONLY_RECORD_PERMISSIONS = ['can_share', 'edit_and_share', 'change_owner'];

/**
 * Folder permissions that are permanent-only (no duration support)
 */
const PERMANENT_ONLY_FOLDER_PERMISSIONS = ['manage_users', 'manage_all'];

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
      searchType: 'record',
    };
  }
  
  // If action is 'search_folders', extract the data from the card action
  if (requestData.action === 'search_folders') {
    type = 'search-folder';
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
      folderName: requestData.folderName,
      searchType: 'folder',
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
  const itemType = (searchType === 'search-folder' || searchType === 'search_folders') ? 'folder' : 'record';

  if (searchType === 'search-record' || searchType === 'search_records') {
    if (searchQuery && searchQuery.trim()) {
      console.log('[TaskModule] Auto-executing record search on modal open for:', searchQuery);
      return await handleSearchAction(searchQuery, searchApprovalId, contextData, true, 'view_only', 'record');
    } else {
      return buildSearchModal('', searchApprovalId, false, contextData, [], true, 'view_only', null, 'record');
    }
  }
  
  if (searchType === 'search-folder' || searchType === 'search_folders') {
    if (searchQuery && searchQuery.trim()) {
      console.log('[TaskModule] Auto-executing folder search on modal open for:', searchQuery);
      return await handleSearchAction(searchQuery, searchApprovalId, contextData, true, 'no_permissions', 'folder');
    } else {
      return buildSearchModal('', searchApprovalId, false, contextData, [], true, 'no_permissions', null, 'folder');
    }
  }

  return buildSearchModal(searchQuery, searchApprovalId, false, contextData, [], true, 'view_only', null, 'record');
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
  
  const { action, searchQuery, selectedUid, approvalId, permission, duration, searchType } = submitData;
  
  // Determine search type from context or submit data
  const itemType = searchType || approvalContext?.searchType || 'record';

  console.log('[TaskModule] Submit:', { 
    action, 
    searchQuery, 
    selectedUid, 
    approvalId,
    permission,
    duration,
    searchType: itemType,
    hasApprovalContext: !!approvalContext,
    hasCachedResults: !!cachedResults
  });

  if (action === 'search' || action === 'refine_search') {
    // Determine showDuration based on current permission and item type
    const permanentPerms = itemType === 'folder' ? PERMANENT_ONLY_FOLDER_PERMISSIONS : PERMANENT_ONLY_RECORD_PERMISSIONS;
    const defaultPerm = itemType === 'folder' ? 'no_permissions' : 'view_only';
    const showDuration = !permanentPerms.includes(permission || defaultPerm);
    return await handleSearchAction(searchQuery, approvalId, approvalContext || {}, showDuration, permission || defaultPerm, itemType);
  } else if (action === 'select_and_approve') {
    return await handleSelectAndApprove(context, selectedUid, approvalId, permission, duration, approvalContext || {}, itemType);
  } else if (action === 'create_new_record') {
    // Show create record modal
    return handleCreateRecordFetch(searchQuery, approvalContext || {});
  } else if (action === 'submit_create_record') {
    // Handle record creation
    return await handleCreateRecordSubmit(context, submitData, approvalContext || {});
  } else if (action === 'cancel_create_record') {
    // Just close the modal
    return null;
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
 * @param {string} searchType - 'record' or 'folder'
 */
function buildSearchModal(query, approvalId, isLoading = false, approvalContext = {}, results = [], showDuration = true, currentPermission = 'view_only', selectedUid = null, searchType = 'record') {
  
  const isFolder = searchType === 'folder';
  const itemLabel = isFolder ? 'folder' : 'record';
  const itemLabelCap = isFolder ? 'Folder' : 'Record';
  const itemLabelPlural = isFolder ? 'Folders' : 'Records';
  const permissions = isFolder ? FOLDER_PERMISSIONS : RECORD_PERMISSIONS;
  const defaultPermission = isFolder ? 'no_permissions' : 'view_only';
  
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
    placeholder: `Enter ${itemLabel} name or UID...`,
    value: query || '',
  });
  
  body.push({
    type: 'TextBlock',
    text: 'Modify the search term and click the Refine button below.',
    wrap: true,
    isSubtle: true,
    size: 'Small',
  });
  
  // Inline buttons with spacing: Refine Search + Create New Record (only for records)
  if (isFolder) {
    // Folders only get Refine Search button
    body.push({
      type: 'ActionSet',
      actions: [
        {
          type: 'Action.Submit',
          title: '🔍 Refine Search',
          data: { action: 'refine_search', searchType: 'folder' },
        },
      ],
    });
  } else {
    // Records get both buttons
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
                  title: '🔍 Refine Search',
                  data: { action: 'refine_search', searchType: 'record' },
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
                  data: { action: 'create_new_record', searchType: 'record' },
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
  }
  
  // Results section
  if (isLoading) {
    body.push({
      type: 'TextBlock',
      text: '🔄 Searching...',
      color: 'Accent',
      wrap: true,
    });
  } else if (results.length > 0) {
    // Check if this is a newly created record
    const newlyCreatedUid = approvalContext?.newlyCreatedUid;
    const newlyCreatedTitle = approvalContext?.newlyCreatedTitle;
    
    if (newlyCreatedUid && results.length === 1 && results[0].uid === newlyCreatedUid) {
      body.push({
        type: 'TextBlock',
        text: `✅ New record "${newlyCreatedTitle}" created successfully!`,
        color: 'Good',
        wrap: true,
        weight: 'Bolder',
      });
      body.push({
        type: 'TextBlock',
        text: 'Select the record below and click "Approve Access" to grant access to the requester.',
        wrap: true,
        isSubtle: true,
        size: 'Small',
      });
    } else {
      body.push({
        type: 'TextBlock',
        text: `Showing ${results.length} result(s) for: **${query}**`,
        wrap: true,
      });
    }
    
    // Item selection with required indicator
    body.push({
      type: 'TextBlock',
      text: `Select ${itemLabel}: *`,
      weight: 'Bolder',
    });
    
    body.push({
      type: 'Input.ChoiceSet',
      id: 'selectedUid',
      style: 'expanded',
      choices: results.map(r => ({ 
        title: `${r.title || r.name} (${r.uid})`, 
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
      choices: permissions,
      value: currentPermission || defaultPermission,
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
      text: `❌ No ${itemLabelPlural.toLowerCase()} found matching "${query}"`,
      color: 'Attention',
      wrap: true,
    });
  }
  
  // Hidden context fields
  body.push({
    type: 'Input.Text',
    id: 'approvalContext',
    isVisible: false,
    value: JSON.stringify(approvalContext),
  });
  
  body.push({
    type: 'Input.Text',
    id: 'searchType',
    isVisible: false,
    value: searchType,
  });
  
  // Bottom action - only Approve Access (when results exist)
  const actions = [];
  if (results && results.length > 0) {
    actions.push({
      type: 'Action.Submit',
      title: 'Approve Access',
      style: 'positive',
      data: { action: 'select_and_approve', searchType: searchType },
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
        title: `Search ${itemLabelPlural}`,
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
 * @param {string} searchType - 'record' or 'folder'
 */
async function handleSearchAction(query, approvalId, approvalContext, showDuration = true, currentPermission = 'view_only', searchType = 'record') {
  const isFolder = searchType === 'folder';
  const defaultPermission = isFolder ? 'no_permissions' : 'view_only';
  
  if (!query || !query.trim()) {
    return buildSearchModal('', approvalId, false, approvalContext, [], showDuration, currentPermission || defaultPermission, null, searchType);
  }

  console.log(`[TaskModule] Executing ${searchType} search for:`, query);

  let results = [];
  
  if (isFolder) {
    // Folder search
    let folder = await keeperClient.getFolderByUid(query.trim());
    
    if (folder) {
      results = [{ uid: folder.uid, title: folder.name || folder.uid, name: folder.name }];
      console.log('[TaskModule] Found folder by UID:', folder.uid);
    } else {
      const searchResults = await keeperClient.searchFolders(query, 10);
      results = searchResults.map(f => ({ uid: f.uid, title: f.name || f.uid, name: f.name }));
      console.log('[TaskModule] Found', results.length, 'folders by search');
    }
  } else {
    // Record search
    let record = await keeperClient.getRecordByUid(query.trim());
    
    if (record) {
      results = [{ uid: record.uid, title: record.title || record.uid }];
      console.log('[TaskModule] Found record by UID:', record.uid);
    } else {
      const searchResults = await keeperClient.searchRecords(query, 10);
      results = searchResults.map(r => ({ uid: r.uid, title: r.title || r.uid }));
      console.log('[TaskModule] Found', results.length, 'records by search');
    }
  }

  return buildSearchModal(query, approvalId, false, approvalContext, results, showDuration, currentPermission || defaultPermission, null, searchType);
}

/**
 * Handle select and approve - grant access and close modal
 * @param {string} searchType - 'record' or 'folder'
 */
async function handleSelectAndApprove(context, selectedUid, approvalId, permission, duration, approvalContext, searchType = 'record') {
  const isFolder = searchType === 'folder';
  const itemLabel = isFolder ? 'folder' : 'record';
  const itemLabelCap = isFolder ? 'Folder' : 'Record';
  
  console.log(`[TaskModule] Approving ${itemLabel}:`, { selectedUid, approvalId, permission, duration });

  if (!selectedUid) {
    return {
      task: {
        type: 'message',
        value: `❌ Please select a ${itemLabel} first.`,
      },
    };
  }

  // Force permanent duration for permanent-only permissions
  const permanentPerms = isFolder ? PERMANENT_ONLY_FOLDER_PERMISSIONS : PERMANENT_ONLY_RECORD_PERMISSIONS;
  if (permanentPerms.includes(permission)) {
    console.log(`[TaskModule] Permission "${permission}" is permanent-only, forcing duration to permanent`);
    duration = 'permanent';
  }

  // Fetch the item
  let item = null;
  let itemTitle = selectedUid;
  
  if (isFolder) {
    item = await keeperClient.getFolderByUid(selectedUid);
    if (item) {
      itemTitle = item.name || selectedUid;
    }
  } else {
    item = await keeperClient.getRecordByUid(selectedUid);
    if (item) {
      itemTitle = item.title || selectedUid;
    }
  }
  
  if (!item) {
    return {
      task: {
        type: 'message',
        value: `❌ ${itemLabelCap} not found: ${selectedUid}`,
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

  console.log(`[TaskModule] Granting ${itemLabel} access:`, {
    uid: selectedUid,
    requesterEmail,
    permission,
    durationSeconds,
  });

  // Grant access based on item type
  let result;
  if (isFolder) {
    result = await keeperClient.grantFolderAccess(
      selectedUid,
      requesterEmail,
      permission,
      durationSeconds
    );
  } else {
    result = await keeperClient.grantRecordAccess(
      selectedUid,
      requesterEmail,
      permission,
      durationSeconds
    );
  }

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
    let updatedCard;
    if (isFolder) {
      updatedCard = cards.buildFolderApprovalCardWithStatus({
        approvalId: approvalContext.approvalId,
        requesterName: approvalContext.requesterName,
        requesterEmail: requesterEmail,
        folderName: itemTitle,
        justification: approvalContext.justification,
        status: 'approved',
        approverName: approverName,
        permission: permission,
        duration: durationDisplay,
        expiresAt: expiresAtFormatted,
      });
    } else {
      updatedCard = cards.buildRecordApprovalCardWithStatus({
        approvalId: approvalContext.approvalId,
        requesterName: approvalContext.requesterName,
        requesterEmail: requesterEmail,
        recordTitle: itemTitle,
        justification: approvalContext.justification,
        status: 'approved',
        approverName: approverName,
        permission: permission,
        duration: durationDisplay,
        expiresAt: expiresAtFormatted,
      });
    }
    
    // Try to update the existing approval card in-place
    const activityId = getApprovalActivityId(approvalContext.approvalId);
    let cardUpdated = false;
    
    if (activityId) {
      console.log(`[TaskModule] Attempting to update approval card in-place (approvalId: ${approvalContext.approvalId}, activityId: ${activityId})`);
      try {
        const channelService = getChannelService();
        
        if (channelService) {
          // Use the channel service's update method (takes activityId and card)
          const updateSuccess = await channelService.updateApprovalCard(activityId, updatedCard);
          
          if (updateSuccess) {
            console.log(`[TaskModule] Successfully updated ${itemLabel} approval card in-place`);
            cardUpdated = true;
          }
        }
      } catch (updateError) {
        console.error('[TaskModule] Error updating card in-place:', updateError.message);
      }
    } else {
      console.log('[TaskModule] No activity ID found for approval, will send new card');
    }
    
    // Fallback: Send new card if update failed
    if (!cardUpdated) {
      try {
        await context.send({
          type: 'message',
          attachments: [{
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: updatedCard,
          }],
        });
        console.log(`[TaskModule] Sent new ${itemLabel} approval status card (fallback)`);
      } catch (sendError) {
        console.error('[TaskModule] Error sending approval status:', sendError.message);
      }
    }
    
    // Send DM notification to the requester
    const requesterId = approvalContext.requesterId;
    if (requesterId) {
      try {
        const channelService = getChannelService();
        if (channelService) {
          const notificationCard = cards.buildRequesterNotificationCard({
            approved: true,
            recordTitle: itemTitle,
            itemType: itemLabel,
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
        value: `❌ Failed to grant ${itemLabel} access: ${result.error || 'Unknown error'}`,
      },
    };
  }
}

/**
 * Handle create record fetch - show the create record modal
 */
function handleCreateRecordFetch(searchQuery, approvalContext) {
  console.log('[TaskModule] Opening create record modal for:', { searchQuery, requester: approvalContext.requesterName });
  
  const createRecordCard = cards.buildCreateRecordModal(approvalContext, searchQuery);
  
  return {
    task: {
      type: 'continue',
      value: {
        title: 'New Record Creation',
        height: 600,
        width: 450,
        card: {
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: createRecordCard,
        },
      },
    },
  };
}

/**
 * Handle create record submit - create the record and return to search with pre-selected result
 */
async function handleCreateRecordSubmit(context, submitData, approvalContext) {
  const { recordTitle, recordLogin, recordPassword, recordUrl, recordNotes } = submitData;
  
  console.log('[TaskModule] Creating record:', { title: recordTitle, login: recordLogin });
  
  // Validate required fields
  if (!recordTitle || !recordTitle.trim()) {
    return {
      task: {
        type: 'message',
        value: '❌ Title is required',
      },
    };
  }
  
  if (!recordLogin || !recordLogin.trim()) {
    return {
      task: {
        type: 'message',
        value: '❌ Login is required',
      },
    };
  }
  
  // Determine if we should generate password
  const generatePassword = !recordPassword || recordPassword.trim() === '' || recordPassword === '$GEN';
  const passwordToUse = generatePassword ? '$GEN' : recordPassword;
  
  // Create the record
  const result = await keeperClient.createRecord({
    title: recordTitle.trim(),
    login: recordLogin.trim(),
    password: passwordToUse,
    url: recordUrl?.trim() || null,
    notes: recordNotes?.trim() || null,
    generatePassword: generatePassword,
  });
  
  if (!result.success) {
    return {
      task: {
        type: 'message',
        value: `❌ Failed to create record: ${result.error || 'Unknown error'}`,
      },
    };
  }
  
  console.log('[TaskModule] Record created successfully:', result.recordUid);
  
  // Build the search modal with the newly created record pre-selected
  const newRecord = {
    uid: result.recordUid,
    title: recordTitle.trim(),
  };
  
  // Update approval context with newly created record info
  const updatedContext = {
    ...approvalContext,
    newlyCreatedUid: result.recordUid,
    newlyCreatedTitle: recordTitle.trim(),
  };
  
  // Return to search modal with the new record shown and pre-selected
  return buildSearchModal(
    recordTitle.trim(),
    approvalContext.approvalId,
    false,
    updatedContext,
    [newRecord],  // Show only the newly created record
    true,  // showDuration
    'view_only',
    result.recordUid,  // Pre-select the new record
    'record'
  );
}

module.exports = {
  handleTaskFetch,
  handleTaskSubmit,
  handleCreateRecordFetch,
  handleCreateRecordSubmit,
};
