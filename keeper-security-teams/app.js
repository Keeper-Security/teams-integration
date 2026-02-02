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
    
    // Handle search_records action
    if (action === 'search_records') {
      console.log('[Keeper Bot] Search records action');
      const taskModuleResponse = await handlers.handleSearchRecordsAction(context, data);
      if (taskModuleResponse) {
        return taskModuleResponse;
      }
      return;
    }
    
    // Handle search_folders action
    if (action === 'search_folders') {
      console.log('[Keeper Bot] Search folders action');
      const taskModuleResponse = await handlers.handleSearchFoldersAction(context, data);
      if (taskModuleResponse) {
        return taskModuleResponse;
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
    
    if (action?.startsWith('approve_') || action?.startsWith('deny_')) {
      try {
        console.log(`[Keeper Bot] Processing ${action} from adaptiveCard/action (Universal Action)`);
        
        // Call the handler and get the updated card
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

  // Handle search_records action without msteams (fallback)
  if (action === 'search_records') {
    console.log('[Keeper Bot] Search records action without msteams, handling manually');
    const taskModuleResponse = await handlers.handleSearchRecordsAction(context, data);
    if (taskModuleResponse) {
      return taskModuleResponse;
    }
    return;
  }
  
  // Handle search_folders action without msteams (fallback)
  if (action === 'search_folders') {
    console.log('[Keeper Bot] Search folders action without msteams, handling manually');
    const taskModuleResponse = await handlers.handleSearchFoldersAction(context, data);
    if (taskModuleResponse) {
      return taskModuleResponse;
    }
    return;
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
        await handlers.routePedmAction(context, data);
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
    console.log('[Keeper Bot] PEDM poller started');
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

// Extend the start method to initialize pollers
const originalStart = app.start.bind(app);
app.start = async (...args) => {
  const result = await originalStart(...args);
  
  // Start background pollers after app is running
  setTimeout(() => {
    startPollers();
  }, 2000); // Delay to ensure app is fully initialized
  
  return result;
};

module.exports = app;
