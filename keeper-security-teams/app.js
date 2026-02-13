/**
 * Keeper Teams Bot Application
 * 
 * Main bot application built on Microsoft Teams SDK.
 * Handles commands, card actions, and integrates with Keeper Commander.
 */

const { stripMentionsText } = require("@microsoft/teams.api");
const { App } = require("@microsoft/teams.apps");
const { LocalStorage } = require("@microsoft/teams.common");
const { ManagedIdentityCredential } = require("@azure/identity");

const config = require("./config");
const handlers = require("./handlers");
const { PedmPoller, DevicePoller } = require("./background");
const { initializeChannelService, getChannelService, isApprovalsChannel } = require("./services");

// Create storage for conversation state
const storage = new LocalStorage();

// ==================== Authentication Setup ====================

const createTokenFactory = () => {
  return async (scope, tenantId) => {
    const managedIdentityCredential = new ManagedIdentityCredential({
      clientId: process.env.CLIENT_ID,
    });
    const scopes = Array.isArray(scope) ? scope : [scope];
    const tokenResponse = await managedIdentityCredential.getToken(scopes, {
      tenantId: tenantId,
    });
    return tokenResponse.token;
  };
};

const tokenCredentials = {
  clientId: process.env.CLIENT_ID || "",
  token: createTokenFactory(),
};

const credentialOptions =
  config.MicrosoftAppType === "UserAssignedMsi" ? { ...tokenCredentials } : undefined;

// ==================== Create Teams App ====================

const app = new App({
  ...credentialOptions,
  storage,
});

// ==================== Channel Service Initialization ====================

// Initialize the channel service for proactive messaging
const channelService = initializeChannelService(app);

// ==================== Conversation State Management ====================

const getConversationState = (conversationId) => {
  let state = storage.get(conversationId);
  if (!state) {
    state = { 
      count: 0,
      pendingRequests: {}, // Track pending approval requests
    };
    storage.set(conversationId, state);
  }
  return state;
};

// ==================== Message Handler ====================

app.on("message", async (context) => {
  const activity = context.activity;
  
  // Check if this is a card action (Action.Submit sends data in activity.value)
  if (activity.value && activity.value.action) {
    const data = activity.value;
    const action = data.action;
    
    console.log(`[Keeper Bot] Card action in message:`, JSON.stringify(data));
    
    // Handle approval/denial actions
    if (action?.startsWith('approve_') || action?.startsWith('deny_')) {
      try {
        console.log(`[Keeper Bot] Routing ${action}`);
        await handlers.routeApprovalAction(context, data);
        console.log(`[Keeper Bot] Action ${action} completed`);
      } catch (error) {
        console.error(`[Keeper Bot] Error handling ${action}:`, error);
        await context.send(`❌ Error: ${error.message}`);
      }
      return;
    }
  }
  
  const text = stripMentionsText(activity);

  // Skip if no text
  if (!text) {
    return;
  }

  console.log(`[Keeper Bot] Received message: "${text}"`);

  // ==================== Capture Conversation References ====================
  
  // Capture reference for approvals channel (needed for proactive messaging)
  if (isApprovalsChannel(activity)) {
    channelService.captureConversationReference(context, 'approvals');
    console.log('[Keeper Bot] Captured approvals channel reference');
  }
  
  // Capture user reference for DMs (1:1 chats)
  channelService.captureUserReference(context);

  // ==================== Built-in Debug Commands ====================

  if (text === "/reset") {
    storage.delete(activity.conversation.id);
    await context.send("Ok I've deleted the current conversation state.");
    return;
  }

  if (text === "/count") {
    const state = getConversationState(activity.conversation.id);
    await context.send(`The count is ${state.count}`);
    return;
  }

  if (text === "/diag") {
    await context.send(JSON.stringify(activity, null, 2));
    return;
  }

  if (text === "/state") {
    const state = getConversationState(activity.conversation.id);
    await context.send(JSON.stringify(state, null, 2));
    return;
  }

  if (text === "/runtime") {
    const runtime = {
      nodeversion: process.version,
      sdkversion: "2.0.0",
      keeperServiceUrl: config.keeper.serviceUrl,
      pedmEnabled: config.pedm.enabled,
      deviceApprovalEnabled: config.deviceApproval.enabled,
    };
    await context.send(JSON.stringify(runtime, null, 2));
    return;
  }

  if (text === "/channel-status") {
    const status = channelService.getStatus();
    const canSend = status.approvalsChannelReady && status.appReady;
    await context.send(
      `**Channel Service Status**\n\n` +
      `• Approvals Channel Configured: ${status.approvalsChannelConfigured ? '✅ Yes' : '❌ No'}\n` +
      `• Approvals Channel ID: \`${status.approvalsChannelId}\`\n` +
      `• Approvals Channel Ready: ${status.approvalsChannelReady ? '✅ Yes' : '❌ No'}\n` +
      `• App Ready: ${status.appReady ? '✅ Yes' : '❌ No'}\n` +
      `• Stored References: ${status.storedReferences.join(', ') || 'None'}\n\n` +
      (canSend
        ? '✅ Approval routing is active. Requests will be sent to the approvals channel.'
        : '⚠️ To enable approval routing, send a message in the approvals channel first.')
    );
    return;
  }

  // ==================== Keeper Commands ====================

  // Try to route as a Keeper command
  const handled = await handlers.routeCommand(context, text);
  
  if (handled) {
    // Command was handled by Keeper handlers
    const state = getConversationState(activity.conversation.id);
    state.count++;
    return;
  }

  // ==================== Default Response ====================

  // If no command matched, show help or echo
  if (text.toLowerCase() === 'help' || text === '?') {
    await handlers.handleHelp(context);
  } else {
    // Default echo behavior (can be removed in production)
  const state = getConversationState(activity.conversation.id);
  state.count++;
    
    await context.send(
      `I didn't recognize that command. Type **help** to see available commands.\n\n` +
      `You said: "${text}"`
    );
  }
});

// ==================== Conversation Update Handler ====================

// Handle bot being added to a team/channel (auto-capture approvals channel)
app.on("conversationUpdate", async (context) => {
  const activity = context.activity;
  
  // Check if bot was added to the conversation
  const botAdded = activity.membersAdded?.some(
    member => member.id === activity.recipient?.id
  );
  
  if (botAdded) {
    console.log('[Keeper Bot] Bot added to conversation');
    
    // Check if this is the approvals channel - auto-capture reference
    if (isApprovalsChannel(activity)) {
      channelService.captureConversationReference(context, 'approvals');
      console.log('[Keeper Bot] Auto-captured approvals channel reference on install');
      
      // Send welcome message to approvals channel
      await context.send('✅ **Keeper Security Bot** is now active in this approvals channel.\n\nAccess requests will appear here for review.');
    }
    
    // Capture user reference for DMs (1:1 chats)
    channelService.captureUserReference(context);
  }
});

// ==================== Invoke Handler ====================

// Handle invoke activities (for task modules)
// When msteams: { type: 'task/fetch' } is used, Teams sends an invoke activity
app.on("invoke", async (context) => {
  const activity = context.activity;
  const invokeName = activity.name;
  
  console.log('[Keeper Bot] Invoke activity:', invokeName, JSON.stringify(activity.value));
  
  // Handle task/fetch (opening task module)
  if (invokeName === 'task/fetch') {
    try {
      const { handleTaskFetch } = require('./handlers/taskModuleHandler');
      // Pass the full activity - handleTaskFetch will extract the data
      const response = await handleTaskFetch(context, activity);
      console.log('[Keeper Bot] Task module response type:', typeof response);
      console.log('[Keeper Bot] Task module response keys:', response ? Object.keys(response) : 'null');
      if (response && response.task) {
        console.log('[Keeper Bot] Task type:', response.task.type);
        console.log('[Keeper Bot] Task value keys:', response.task.value ? Object.keys(response.task.value) : 'null');
      }
      console.log('[Keeper Bot] Full task module response:', JSON.stringify(response, null, 2));
      return response;
    } catch (error) {
      console.error('[Keeper Bot] Error handling task/fetch:', error);
      console.error('[Keeper Bot] Error stack:', error.stack);
      return {
        task: {
          type: 'message',
          value: 'Error opening task module: ' + error.message,
        },
      };
    }
  }
  
  // Handle task/submit (submitting task module)
  if (invokeName === 'task/submit') {
    try {
      const { handleTaskSubmit } = require('./handlers/taskModuleHandler');
      const response = await handleTaskSubmit(context, activity);
      console.log('[Keeper Bot] Task module submit response:', JSON.stringify(response).substring(0, 200));
      return response;
    } catch (error) {
      console.error('[Keeper Bot] Error handling task/submit:', error);
      return {
        task: {
          type: 'message',
          value: 'Error processing task module: ' + error.message,
        },
      };
    }
  }
  
  // Handle adaptiveCard/action (Action.Execute from Adaptive Cards)
  if (invokeName === 'adaptiveCard/action') {
    console.log('[Keeper Bot] adaptiveCard/action invoke received');
    const verb = activity.value?.action?.verb;
    const data = activity.value?.action?.data || activity.value?.data || activity.value || {};
    const action = data.action || verb;
    
    console.log('[Keeper Bot] Action verb:', verb);
    console.log('[Keeper Bot] Action data:', JSON.stringify(data));
    
    // Handle refresh action - return the correct card based on approval status
    if (verb === 'refreshApprovalCard') {
      try {
        console.log('[Keeper Bot] Processing refresh for approval:', data.approvalId);
        const { handleRefreshApprovalCard } = require('./handlers/approvalHandler');
        const updatedCard = await handleRefreshApprovalCard(data);
        
        if (updatedCard) {
          console.log('[Keeper Bot] Returning refreshed card for approval:', data.approvalId);
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: updatedCard,
          };
        }
        
        // No status change, return empty to keep original card
        console.log('[Keeper Bot] No status change, keeping original card');
        return { statusCode: 200 };
      } catch (error) {
        console.error('[Keeper Bot] Error refreshing approval card:', error);
        return { statusCode: 200 }; // Return 200 to avoid error display
      }
    }
    
    // Handle inline lookup actions (search from the card itself)
    if (verb === 'lookup_record' || verb === 'lookup_folder' || verb === 'lookup_share') {
      try {
        const searchQuery = activity.value?.action?.data?.searchQuery || 
                            activity.value?.data?.searchQuery ||
                            data.searchQuery || '';
        console.log(`[Keeper Bot] Processing ${verb} for query:`, searchQuery);
        const { handleInlineLookup } = require('./handlers/approvalHandler');
        const resultCard = await handleInlineLookup(verb, data, searchQuery);
        
        return {
          statusCode: 200,
          type: 'application/vnd.microsoft.card.adaptive',
          value: resultCard,
        };
      } catch (error) {
        console.error(`[Keeper Bot] Error handling ${verb}:`, error);
        return { statusCode: 500, body: error.message };
      }
    }
    
    // Handle reset card actions (return to original approval card)
    if (verb === 'reset_record_card' || verb === 'reset_folder_card' || verb === 'reset_share_card') {
      try {
        console.log(`[Keeper Bot] Processing ${verb} - resetting to original card`);
        const { handleResetCard } = require('./handlers/approvalHandler');
        const originalCard = handleResetCard(verb, data);
        
        return {
          statusCode: 200,
          type: 'application/vnd.microsoft.card.adaptive',
          value: originalCard,
        };
      } catch (error) {
        console.error(`[Keeper Bot] Error handling ${verb}:`, error);
        return { statusCode: 500, body: error.message };
      }
    }
    
    // Handle show_create_form (show inline create record form)
    if (verb === 'show_create_form') {
      try {
        console.log(`[Keeper Bot] Processing show_create_form`);
        const { handleShowCreateForm } = require('./handlers/approvalHandler');
        const createFormCard = handleShowCreateForm(data);
        
        return {
          statusCode: 200,
          type: 'application/vnd.microsoft.card.adaptive',
          value: createFormCard,
        };
      } catch (error) {
        console.error(`[Keeper Bot] Error handling show_create_form:`, error);
        return { statusCode: 500, body: error.message };
      }
    }
    
    // Handle submit_create_record (create record and show with approval options)
    // Uses fire-and-forget pattern: return "Processing" card immediately, then update via proactive messaging
    if (verb === 'submit_create_record') {
      try {
        console.log(`[Keeper Bot] Processing submit_create_record - START`);
        
        // Extract form data
        const recordTitle = activity.value?.action?.data?.recordTitle || data.recordTitle || '';
        const recordLogin = activity.value?.action?.data?.recordLogin || data.recordLogin || '';
        const recordPassword = activity.value?.action?.data?.recordPassword || data.recordPassword || '';
        const recordUrl = activity.value?.action?.data?.recordUrl || data.recordUrl || '';
        const recordNotes = activity.value?.action?.data?.recordNotes || data.recordNotes || '';
        
        // Extract self-destruct options
        const selfDestructToggle = activity.value?.action?.data?.selfDestruct || data.selfDestruct || 'false';
        const selfDestructEnabled = selfDestructToggle === 'true' || selfDestructToggle === true;
        const selfDestructDuration = selfDestructEnabled 
          ? (activity.value?.action?.data?.selfDestructDuration || data.selfDestructDuration || '24h')
          : null;
        
        // Validation - return form with error message if validation fails
        if (!recordTitle?.trim() || !recordLogin?.trim()) {
          const errors = [];
          if (!recordTitle?.trim()) errors.push('Title is required');
          if (!recordLogin?.trim()) errors.push('Login is required');
          const errorMessage = errors.join('. ');
          
          console.log(`[Keeper Bot] Validation failed: ${errorMessage}`);
          
          const cards = require('./cards');
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: cards.buildRecordCreationCard({
              ...data,
              error: errorMessage,
              recordTitle: recordTitle,
              recordLogin: recordLogin,
              recordPassword: recordPassword,
              recordUrl: recordUrl,
              recordNotes: recordNotes,
            }),
          };
        }
        
        // Store context for proactive messaging
        const conversationRef = {
          serviceUrl: activity.serviceUrl,
          channelId: activity.channelId,
          conversation: activity.conversation,
        };
        const activityId = activity.replyToId || activity.id;
        
        // Fire off async record creation (don't await)
        // Use channelService for proactive messaging since context won't be available later
        const channelService = getChannelService();
        
        (async () => {
          try {
            console.log(`[Keeper Bot] Background: Starting record creation${selfDestructEnabled ? ` (self-destruct: ${selfDestructDuration})` : ''}`);
            const keeperClient = require('./services/keeperClient');
            const generatePassword = !recordPassword || recordPassword.trim() === '' || recordPassword === '$GEN';
            
            const result = await keeperClient.createRecord({
              title: recordTitle.trim(),
              login: recordLogin.trim(),
              password: generatePassword ? '$GEN' : recordPassword,
              url: recordUrl?.trim() || null,
              notes: recordNotes?.trim() || null,
              generatePassword: generatePassword,
              selfDestructDuration: selfDestructDuration,
            });
            
            console.log(`[Keeper Bot] Background: Record creation complete`, result.success ? `UID: ${result.recordUid}` : `Error: ${result.error}`);
            
            // Build the result card with permission and duration dropdowns
            const { RECORD_PERMISSIONS, DURATION_OPTIONS } = require('./cards/constants');
            
            let resultCard;
            if (result.success) {
              // Build body elements
              const bodyElements = [
                { type: 'TextBlock', text: 'Record Created Successfully!', weight: 'Bolder', size: 'Large' },
                { type: 'TextBlock', text: `Requester: ${data.requesterName || 'Unknown'}`, wrap: true },
                { type: 'TextBlock', text: `Justification: ${data.justification || 'N/A'}`, wrap: true, isSubtle: true },
                { 
                  type: 'Container', 
                  style: 'good', 
                  spacing: 'Medium',
                  items: [
                    { type: 'TextBlock', text: `Record: ${recordTitle.trim()}`, wrap: true, weight: 'Bolder' },
                    { type: 'TextBlock', text: `UID: ${result.recordUid}`, size: 'Small', isSubtle: true },
                  ]
                },
              ];
              
              // Add self-destruct notice if enabled
              if (selfDestructEnabled) {
                const durationLabels = { '1h': '1 hour', '24h': '24 hours', '7d': '7 days', '30d': '30 days', '90d': '90 days' };
                const durationLabel = durationLabels[selfDestructDuration] || selfDestructDuration;
                bodyElements.push({
                  type: 'Container',
                  style: 'attention',
                  spacing: 'Medium',
                  items: [
                    { type: 'TextBlock', text: 'Self-Destruct Enabled', weight: 'Bolder', wrap: true },
                    { type: 'TextBlock', text: `This record will auto-delete after ${durationLabel}`, size: 'Small', wrap: true },
                  ]
                });
              }
              
              // Add permission/duration selectors
              bodyElements.push(
                { type: 'TextBlock', text: 'Permission Level', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
                { type: 'Input.ChoiceSet', id: 'permission', value: 'view_only', choices: RECORD_PERMISSIONS },
                { type: 'TextBlock', text: 'Duration', weight: 'Bolder', size: 'Medium', spacing: 'Medium' },
                { type: 'Input.ChoiceSet', id: 'duration', value: '1h', choices: DURATION_OPTIONS }
              );
              
              resultCard = {
                type: 'AdaptiveCard',
                '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
                version: '1.4',
                body: bodyElements,
                actions: [
                  {
                    type: 'Action.Execute',
                    title: 'Approve',
                    style: 'positive',
                    verb: 'approve_record',
                    data: { 
                      action: 'approve_record', 
                      approvalId: data.approvalId || '', 
                      recordUid: result.recordUid, 
                      recordTitle: recordTitle.trim(), 
                      requesterId: data.requesterId || '', 
                      requesterEmail: data.requesterEmail || '', 
                      requesterName: data.requesterName || '', 
                      justification: data.justification || '',
                      selfDestruct: selfDestructEnabled,
                      selfDestructDuration: selfDestructDuration,
                    },
                  },
                  {
                    type: 'Action.Execute',
                    title: 'Deny',
                    style: 'destructive',
                    verb: 'deny_record',
                    data: { 
                      action: 'deny_record', 
                      approvalId: data.approvalId || '', 
                      recordUid: result.recordUid, 
                      recordTitle: recordTitle.trim(), 
                      requesterId: data.requesterId || '', 
                      requesterEmail: data.requesterEmail || '', 
                      requesterName: data.requesterName || '', 
                      justification: data.justification || '',
                    },
                  },
                ],
              };
            } else {
              resultCard = {
                type: 'AdaptiveCard',
                '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
                version: '1.4',
                body: [
                  { type: 'TextBlock', text: 'Error Creating Record', weight: 'Bolder', size: 'Large', color: 'Attention' },
                  { type: 'TextBlock', text: result.error || 'Unknown error occurred', wrap: true },
                ],
                actions: [
                  { type: 'Action.Execute', title: 'Try Again', verb: 'show_create_form', data: data },
                ],
              };
            }
            
            // Send new message with result using channelService
            if (channelService && channelService.isApprovalsChannelReady()) {
              try {
                const sent = await channelService.sendApprovalCardViaConnector(
                  resultCard, 
                  data.approvalId || 'create-record',
                  result.success ? `Record "${recordTitle.trim()}" created successfully!` : null
                );
                console.log(`[Keeper Bot] Background: Sent result card via channelService:`, sent.success);
              } catch (sendError) {
                console.error(`[Keeper Bot] Background: Error sending via channelService:`, sendError.message);
              }
            } else {
              console.error(`[Keeper Bot] Background: ChannelService not available for proactive messaging`);
            }
          } catch (bgError) {
            console.error(`[Keeper Bot] Background: Error in record creation:`, bgError.message);
          }
        })();
        
        // Return "Processing" card immediately (within Teams timeout)
        console.log(`[Keeper Bot] Returning processing card immediately`);
        return {
          statusCode: 200,
          type: 'application/vnd.microsoft.card.adaptive',
          value: {
            type: 'AdaptiveCard',
            '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
            version: '1.4',
            body: [
              { type: 'TextBlock', text: 'Creating Record...', weight: 'Bolder', size: 'Large' },
              { type: 'TextBlock', text: `Title: ${recordTitle.trim()}`, wrap: true },
              { type: 'TextBlock', text: 'Please wait. The result will appear below shortly.', wrap: true, isSubtle: true },
            ],
            actions: [],
          },
        };
      } catch (error) {
        console.error(`[Keeper Bot] Error handling submit_create_record:`, error);
        console.error(`[Keeper Bot] Error stack:`, error.stack);
        return { statusCode: 500, body: error.message };
      }
    }
    
    // Handle cancel_create_form (return to search results card)
    if (verb === 'cancel_create_form') {
      try {
        console.log(`[Keeper Bot] Processing cancel_create_form`);
        const { handleCancelCreateForm } = require('./handlers/approvalHandler');
        const searchCard = await handleCancelCreateForm(data);
        
        return {
          statusCode: 200,
          type: 'application/vnd.microsoft.card.adaptive',
          value: searchCard,
        };
      } catch (error) {
        console.error(`[Keeper Bot] Error handling cancel_create_form:`, error);
        return { statusCode: 500, body: error.message };
      }
    }
    
    // Handle approve_selected_record (when multiple records were found)
    if (verb === 'approve_selected_record') {
      try {
        // Extract selected record from input field
        const selectedRecordJson = activity.value?.action?.data?.selectedRecord || 
                                   activity.value?.data?.selectedRecord ||
                                   data.selectedRecord;
        const permission = activity.value?.action?.data?.permission || 
                          activity.value?.data?.permission ||
                          data.permission || 'view_only';
        const duration = activity.value?.action?.data?.duration || 
                        activity.value?.data?.duration ||
                        data.duration || '1h';
        
        if (!selectedRecordJson) {
          console.error('[Keeper Bot] No record selected');
          return { statusCode: 400, body: 'No record selected' };
        }
        
        let selectedRecord;
        try {
          selectedRecord = JSON.parse(selectedRecordJson);
        } catch (e) {
          console.error('[Keeper Bot] Failed to parse selected record:', e);
          return { statusCode: 400, body: 'Invalid record selection' };
        }
        
        console.log(`[Keeper Bot] Approving selected record:`, selectedRecord);
        
        // Build the data for approval
        const approvalData = {
          ...data,
          action: 'approve_record',
          recordUid: selectedRecord.uid,
          recordTitle: selectedRecord.title,
          permission,
          duration,
        };
        
        // Call the existing approval handler
        const result = await handlers.routeApprovalActionWithCardResponse(context, approvalData);
        
        if (result && result.updatedCard) {
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: result.updatedCard,
          };
        }
        
        return { statusCode: 200 };
      } catch (error) {
        console.error('[Keeper Bot] Error handling approve_selected_record:', error);
        return { statusCode: 500, body: error.message };
      }
    }
    
    // Handle approve_selected_folder (when multiple folders were found)
    if (verb === 'approve_selected_folder') {
      try {
        // Extract selected folder from input field
        const selectedFolderJson = activity.value?.action?.data?.selectedFolder || 
                                   activity.value?.data?.selectedFolder ||
                                   data.selectedFolder;
        const permission = activity.value?.action?.data?.permission || 
                          activity.value?.data?.permission ||
                          data.permission || 'no_permissions';
        const duration = activity.value?.action?.data?.duration || 
                        activity.value?.data?.duration ||
                        data.duration || '1h';
        
        if (!selectedFolderJson) {
          console.error('[Keeper Bot] No folder selected');
          return { statusCode: 400, body: 'No folder selected' };
        }
        
        let selectedFolder;
        try {
          selectedFolder = JSON.parse(selectedFolderJson);
        } catch (e) {
          console.error('[Keeper Bot] Failed to parse selected folder:', e);
          return { statusCode: 400, body: 'Invalid folder selection' };
        }
        
        console.log(`[Keeper Bot] Approving selected folder:`, selectedFolder);
        
        // Build the data for approval
        const approvalData = {
          ...data,
          action: 'approve_folder',
          folderUid: selectedFolder.uid,
          folderName: selectedFolder.name,
          permission,
          duration,
        };
        
        // Call the existing approval handler
        const result = await handlers.routeApprovalActionWithCardResponse(context, approvalData);
        
        if (result && result.updatedCard) {
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: result.updatedCard,
          };
        }
        
        return { statusCode: 200 };
      } catch (error) {
        console.error('[Keeper Bot] Error handling approve_selected_folder:', error);
        return { statusCode: 500, body: error.message };
      }
    }
    
    // Handle approve_selected_share (when multiple records were found for one-time share)
    if (verb === 'approve_selected_share') {
      try {
        // Extract selected record from input field
        const selectedRecordJson = activity.value?.action?.data?.selectedRecord || 
                                   activity.value?.data?.selectedRecord ||
                                   data.selectedRecord;
        const duration = activity.value?.action?.data?.duration || 
                        activity.value?.data?.duration ||
                        data.duration || '24h';
        const editable = activity.value?.action?.data?.editable || 
                        activity.value?.data?.editable ||
                        data.editable || 'false';
        
        if (!selectedRecordJson) {
          console.error('[Keeper Bot] No record selected for share');
          return { statusCode: 400, body: 'No record selected' };
        }
        
        let selectedRecord;
        try {
          selectedRecord = JSON.parse(selectedRecordJson);
        } catch (e) {
          console.error('[Keeper Bot] Failed to parse selected record:', e);
          return { statusCode: 400, body: 'Invalid record selection' };
        }
        
        console.log(`[Keeper Bot] Creating one-time share for selected record:`, selectedRecord);
        
        // Build the data for share approval
        const approvalData = {
          ...data,
          action: 'approve_share',
          recordUid: selectedRecord.uid,
          recordTitle: selectedRecord.title,
          duration,
          editable,
        };
        
        // Call the existing approval handler
        const result = await handlers.routeApprovalActionWithCardResponse(context, approvalData);
        
        if (result && result.updatedCard) {
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: result.updatedCard,
          };
        }
        
        return { statusCode: 200 };
      } catch (error) {
        console.error('[Keeper Bot] Error handling approve_selected_share:', error);
        return { statusCode: 500, body: error.message };
      }
    }
    
    if (action?.startsWith('approve_') || action?.startsWith('deny_')) {
      try {
        console.log(`[Keeper Bot] Processing ${action} from adaptiveCard/action (Universal Action)`);
        
        // Check if this is a PEDM action
        if (action.includes('pedm')) {
          // PEDM actions return a card directly for in-place update (like Slack does)
          const resultCard = await handlers.routePedmAction(context, data);
          if (resultCard) {
            console.log(`[Keeper Bot] PEDM action ${action} completed, updating card in-place`);
            return {
              statusCode: 200,
              type: 'application/vnd.microsoft.card.adaptive',
              value: resultCard,
            };
          }
          return { statusCode: 200 };
        }
        
        // Call the handler and get the updated card for other approval types
        const result = await handlers.routeApprovalActionWithCardResponse(context, data);
        
        console.log(`[Keeper Bot] Action ${action} completed, returning updated card`);
        
        // Return the updated card - Teams will automatically update the original card
        if (result && result.updatedCard) {
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: result.updatedCard,
          };
        }
        
        // Fallback: just acknowledge
        return { statusCode: 200 };
      } catch (error) {
        console.error(`[Keeper Bot] Error handling ${action}:`, error);
        return { statusCode: 500, body: error.message };
      }
    }
  }
  
  // Log any unhandled invoke names for debugging
  console.log('[Keeper Bot] Unhandled invoke name:', invokeName);
  
  // Return undefined to let other handlers process
  return undefined;
});

// ==================== Task Module Handler ====================

// Handle task/fetch - Show task module UI (legacy handler, kept for compatibility)
app.on("task/fetch", async (context) => {
  const activity = context.activity;
  
  console.log('[Keeper Bot] Task module fetch (legacy):', JSON.stringify(activity.value));
  
  try {
    const { handleTaskFetch } = require('./handlers/taskModuleHandler');
    const response = await handleTaskFetch(context, activity);
    return response;
  } catch (error) {
    console.error('[Keeper Bot] Error handling task/fetch:', error);
    return {
      task: {
        type: 'message',
        value: 'Error opening task module: ' + error.message,
      },
    };
  }
});

// Handle task/submit - Process task module submission (legacy handler, kept for compatibility)
app.on("task/submit", async (context) => {
  const activity = context.activity;
  
  console.log('[Keeper Bot] Task module submit (legacy):', JSON.stringify(activity.value));
  
  try {
    const { handleTaskSubmit } = require('./handlers/taskModuleHandler');
    const response = await handleTaskSubmit(context, activity);
    return response;
  } catch (error) {
    console.error('[Keeper Bot] Error handling task/submit:', error);
    return {
      task: {
        type: 'message',
        value: 'Error processing task module: ' + error.message,
      },
    };
  }
});

// ==================== Card Action Handler ====================

// Handle Adaptive Card action submissions
app.on("cardAction", async (context) => {
  const activity = context.activity;
  const data = activity.value || {};
  
  console.log(`[Keeper Bot] Card action received:`, JSON.stringify(data));

  const action = data.action;

  if (!action) {
    console.log('[Keeper Bot] No action in card data, skipping');
    return;
  }

  // If this action has msteams.task/fetch, it will be handled by invoke handler
  // Don't intercept it here - let Teams convert it to an invoke activity
  if (data.msteams && data.msteams.type === 'task/fetch') {
    console.log('[Keeper Bot] Task module request detected in cardAction, should be handled by invoke handler');
    // Return undefined to let the invoke handler process it
    // Teams will convert this to an invoke activity automatically
    return undefined;
  }

  // Route to appropriate handler based on action type
  try {
    if (action?.startsWith('approve_') || action?.startsWith('deny_')) {
      console.log(`[Keeper Bot] Routing approval/denial action: ${action}`);
      // Approval/denial actions
      if (action.includes('record') || action.includes('folder') || action.includes('share')) {
        await handlers.routeApprovalAction(context, data);
        console.log(`[Keeper Bot] Action ${action} completed`);
      } else if (action.includes('pedm')) {
        // PEDM actions return a card for in-place update (like Slack does)
        const resultCard = await handlers.routePedmAction(context, data);
        if (resultCard) {
          console.log(`[Keeper Bot] PEDM action ${action} completed, updating card in-place`);
          // Return card for in-place update
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: resultCard,
          };
        }
      } else if (action.includes('device')) {
        await handlers.routeDeviceAction(context, data);
      }
    } else {
      console.log(`[Keeper Bot] Unknown action: ${action}`);
    }
  } catch (error) {
    console.error(`[Keeper Bot] Error handling card action:`, error);
    await context.send(`❌ Error processing action: ${error.message}`);
  }
});

// ==================== Background Pollers ====================

// Initialize pollers (they will only start if enabled in config)
let pedmPoller = null;
let devicePoller = null;

// Start pollers after app is ready
const startPollers = () => {
  if (config.pedm.enabled) {
    pedmPoller = new PedmPoller(app);
    pedmPoller.start();
    console.log('[Keeper Bot] EPM poller started');
  }

  if (config.deviceApproval.enabled) {
    devicePoller = new DevicePoller(app);
    devicePoller.start();
    console.log('[Keeper Bot] Device approval poller started');
  }
};

// Stop pollers on shutdown
const stopPollers = () => {
  if (pedmPoller) pedmPoller.stop();
  if (devicePoller) devicePoller.stop();
};

// Handle graceful shutdown
process.on('SIGINT', () => {
  console.log('[Keeper Bot] Shutting down...');
  stopPollers();
  process.exit(0);
});

process.on('SIGTERM', () => {
  console.log('[Keeper Bot] Shutting down...');
  stopPollers();
  process.exit(0);
});

// ==================== App Ready Hook ====================

// Extend the start method to initialize pollers and check connectivity
const originalStart = app.start.bind(app);
app.start = async (...args) => {
  const result = await originalStart(...args);
  
  console.log('\n' + '='.repeat(60));
  console.log('Starting Keeper Commander Teams Bot');
  console.log('='.repeat(60));
  
  // Check Keeper Service Mode connectivity
  const keeperClient = require('./services/keeperClient');
  const serviceUrl = config.keeper?.serviceUrl || process.env.KEEPER_SERVICE_URL || 'http://localhost:3001/api/v2/';
  
  try {
    const isHealthy = await keeperClient.healthCheck();
    
    if (isHealthy) {
      console.log('[Keeper Bot] ✓ Keeper Service Mode is accessible');
    } else {
      console.warn('[Keeper Bot] ⚠ Cannot reach Keeper Service Mode');
      console.warn(`   URL: ${serviceUrl}`);
      console.warn('   The bot will start but commands may fail.');
    }
  } catch (error) {
    console.warn('[Keeper Bot] ⚠ Cannot reach Keeper Service Mode');
    console.warn(`   URL: ${serviceUrl}`);
    console.warn(`   Error: ${error.message}`);
    console.warn('   The bot will start but commands may fail.');
  }
  
  console.log('='.repeat(60) + '\n');
  
  // Start background pollers after app is running
  setTimeout(() => {
    startPollers();
  }, 2000); // Delay to ensure app is fully initialized
  
  return result;
};

module.exports = app;
