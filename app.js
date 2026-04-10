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

const { getConfig } = require("./config");
const handlers = require("./handlers");
const { EpmPoller, DevicePoller } = require("./background");
const { initializeChannelService, getChannelService, isApprovalsChannel, createLogger } = require("./services");

const log = createLogger('KeeperBot');

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
  getConfig().MicrosoftAppType === "UserAssignedMsi" ? { ...tokenCredentials } : undefined;

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
  
  if (activity.value && activity.value.action) {
    const data = activity.value;
    const action = data.action;
    
    log.debug('Card action in message', data);
    
    if (action?.startsWith('approve_') || action?.startsWith('deny_')) {
      try {
        log.debug(`Routing ${action}`);
        await handlers.routeApprovalAction(context, data);
        log.debug(`Action ${action} completed`);
      } catch (error) {
        log.error(`Error handling ${action}`, error);
        await context.send(`Error: ${error.message}`);
      }
      return;
    }
  }
  
  const text = stripMentionsText(activity);

  // Skip if no text
  if (!text) {
    return;
  }

  log.info(`Received message: "${text}"`);

  // ==================== Capture Conversation References ====================
  log.info('Channel check debug', {
    teamsChannelId: activity.channelData?.teamsChannelId,
    conversationId: activity.conversation?.id,
    configuredId: getConfig().teams?.approvalsChannelId,
  });
  if (isApprovalsChannel(activity)) {
    channelService.captureConversationReference(context, 'approvals');
    log.info('Captured approvals channel reference');
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
    const currentConfig = getConfig();
    const runtime = {
      nodeversion: process.version,
      sdkversion: "2.0.0",
      keeperServiceUrl: currentConfig.keeper.serviceUrl,
      pedmEnabled: currentConfig.pedm.enabled,
      deviceApprovalEnabled: currentConfig.deviceApproval.enabled,
    };
    await context.send(JSON.stringify(runtime, null, 2));
    return;
  }

  if (text === "/channel-status") {
    const status = channelService.getStatus();
    const canSend = status.approvalsChannelReady && status.appReady;
    await context.send(
      `**Channel Service Status**\n\n` +
      `• Approvals Channel Configured: ${status.approvalsChannelConfigured ? 'Yes' : 'No'}\n` +
      `• Approvals Channel ID: \`${status.approvalsChannelId}\`\n` +
      `• Approvals Channel Ready: ${status.approvalsChannelReady ? 'Yes' : 'No'}\n` +
      `• App Ready: ${status.appReady ? 'Yes' : 'No'}\n` +
      `• Stored References: ${status.storedReferences.join(', ') || 'None'}\n\n` +
      (canSend
        ? 'Approval routing is active. Requests will be sent to the approvals channel.'
        : 'To enable approval routing, send a message in the approvals channel first.')
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
    log.info('Bot added to conversation');
    
    if (isApprovalsChannel(activity)) {
      channelService.captureConversationReference(context, 'approvals');
      log.info('Auto-captured approvals channel reference on install');
      
      // Send welcome message to approvals channel
      await context.send('**Keeper Security Bot** is now active in this approvals channel.\n\nAccess requests will appear here for review.');
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
  
  log.debug('Invoke activity', { name: invokeName, value: activity.value });
  
  // Handle task/fetch (opening task module)
  if (invokeName === 'task/fetch') {
    try {
      const { handleTaskFetch } = require('./handlers/taskModuleHandler');
      const response = await handleTaskFetch(context, activity);
      log.debug('Task module response', response);
      return response;
    } catch (error) {
      log.error('Error handling task/fetch', { message: error.message, stack: error.stack });
      return {
        task: {
          type: 'message',
          value: 'Error opening task module: ' + error.message,
        },
      };
    }
  }
  
  if (invokeName === 'task/submit') {
    try {
      const { handleTaskSubmit } = require('./handlers/taskModuleHandler');
      const response = await handleTaskSubmit(context, activity);
      log.debug('Task module submit response');
      return response;
    } catch (error) {
      log.error('Error handling task/submit', error.message);
      return {
        task: {
          type: 'message',
          value: 'Error processing task module: ' + error.message,
        },
      };
    }
  }
  
  if (invokeName === 'adaptiveCard/action') {
    log.debug('adaptiveCard/action invoke received');
    const verb = activity.value?.action?.verb;
    const data = activity.value?.action?.data || activity.value?.data || activity.value || {};
    const action = data.action || verb;
    
    log.debug('Action', { verb, action });
    
    if (verb === 'refreshApprovalCard') {
      try {
        log.debug('Processing refresh for approval', data.approvalId);
        const { handleRefreshApprovalCard } = require('./handlers/approvalHandler');
        const updatedCard = await handleRefreshApprovalCard(data);
        
        if (updatedCard) {
          log.debug('Returning refreshed card for approval', data.approvalId);
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: updatedCard,
          };
        }
        
        log.debug('No status change, keeping original card');
        return { statusCode: 200 };
      } catch (error) {
        log.error('Error refreshing approval card', error);
        return { statusCode: 200 };
      }
    }
    
    if (verb === 'lookup_record' || verb === 'lookup_folder' || verb === 'lookup_share') {
      try {
        const searchQuery = activity.value?.searchQuery ||
                            activity.value?.action?.data?.searchQuery || 
                            activity.value?.data?.searchQuery ||
                            data.searchQuery || '';
        log.debug(`Processing ${verb} for query: ${searchQuery}`);
        const { handleInlineLookup } = require('./handlers/approvalHandler');
        const resultCard = await handleInlineLookup(verb, data, searchQuery);
        
        return {
          statusCode: 200,
          type: 'application/vnd.microsoft.card.adaptive',
          value: resultCard,
        };
      } catch (error) {
        log.error(`Error handling ${verb}`, error);
        return { statusCode: 500, body: error.message };
      }
    }
    
    if (verb === 'reset_record_card' || verb === 'reset_folder_card' || verb === 'reset_share_card') {
      try {
        log.debug(`Processing ${verb} - resetting to original card`);
        const { handleResetCard } = require('./handlers/approvalHandler');
        const originalCard = handleResetCard(verb, data);
        
        return {
          statusCode: 200,
          type: 'application/vnd.microsoft.card.adaptive',
          value: originalCard,
        };
      } catch (error) {
        log.error(`Error handling ${verb}`, error);
        return { statusCode: 500, body: error.message };
      }
    }
    
    if (verb === 'show_create_form') {
      try {
        log.debug('Processing show_create_form');
        const { handleShowCreateForm } = require('./handlers/approvalHandler');
        const createFormCard = handleShowCreateForm(data);
        
        return {
          statusCode: 200,
          type: 'application/vnd.microsoft.card.adaptive',
          value: createFormCard,
        };
      } catch (error) {
        log.error('Error handling show_create_form', error);
        return { statusCode: 500, body: error.message };
      }
    }
    
    if (verb === 'submit_create_record') {
      try {
        log.debug('Processing submit_create_record - START');

        const v = activity.value || {};
        const merged = {
          ...data,
          ...(v.action?.data || {}),
          ...(v.data || {}),
        };

        const isCreateSecretFlow =
          merged.createSecretFlow === true ||
          merged.createSecretFlow === 'true';

        const recordTitle = merged.recordTitle || '';
        const recordLogin = merged.recordLogin || '';
        const recordPassword = merged.recordPassword || '';
        const recordUrl = merged.recordUrl || '';
        const recordNotes = merged.recordNotes || '';
        const targetFolderUidRaw = merged.targetFolderUid;

        let selfDestructEnabled = false;
        let selfDestructDuration = null;
        if (!isCreateSecretFlow) {
          const selfDestructToggle = merged.selfDestruct || 'false';
          selfDestructEnabled = selfDestructToggle === 'true' || selfDestructToggle === true;
          selfDestructDuration = selfDestructEnabled
            ? (merged.selfDestructDuration || '24h')
            : null;
        }

        const cards = require('./cards');

        const buildApprovalValidationCard = (errorMessage) =>
          cards.buildRecordCreationCard({
            approvalId: merged.approvalId,
            requesterName: merged.requesterName,
            requesterId: merged.requesterId,
            requesterEmail: merged.requesterEmail,
            requesterAadObjectId: merged.requesterAadObjectId,
            justification: merged.justification,
            identifier: merged.identifier,
            originalRecordTitle: merged.originalRecordTitle,
            searchQuery: merged.searchQuery,
            createSecretFlow: false,
            error: errorMessage,
            recordTitle,
            recordLogin,
            recordPassword,
            recordUrl,
            recordNotes,
          });

        const pickValidTargetFolderUid = (raw, choices) => {
          if (!choices?.length) return '_default_';
          const allowed = new Set(choices.map((c) => c.value));
          const v = raw != null ? String(raw).trim() : '';
          if (v && allowed.has(v)) return v;
          return '_default_';
        };

        const buildCreateSecretCardFromFolderResult = (folderResult, errorMessage) => {
          const choiceSetChoices = folderResult.choiceSetChoices;
          const sharedFoldersLoadError = folderResult.error;
          const noSharedFoldersForUser = !!folderResult.noSharedFoldersForUser;
          const selected = pickValidTargetFolderUid(targetFolderUidRaw, choiceSetChoices);
          return cards.buildRecordCreationCard({
            approvalId: merged.approvalId,
            requesterName: merged.requesterName,
            requesterId: merged.requesterId,
            requesterEmail: merged.requesterEmail,
            requesterAadObjectId: merged.requesterAadObjectId,
            justification: merged.justification,
            identifier: merged.identifier,
            originalRecordTitle: merged.originalRecordTitle,
            searchQuery: merged.searchQuery,
            createSecretFlow: true,
            sharedFolderChoices: choiceSetChoices,
            sharedFoldersLoadError,
            noSharedFoldersForUser,
            selectedTargetFolderUid: selected,
            ...(errorMessage ? { error: errorMessage } : {}),
            recordTitle,
            recordLogin,
            recordPassword,
            recordUrl,
            recordNotes,
          });
        };

        const buildCreateSecretValidationCard = async (errorMessage) => {
          const keeperClient = require('./services/keeperClient');
          const folderResult = await keeperClient.getSharedFolderChoicesForEmail(merged.requesterEmail || '');
          return buildCreateSecretCardFromFolderResult(folderResult, errorMessage);
        };

        if (isCreateSecretFlow) {
          const folderSelected = targetFolderUidRaw && String(targetFolderUidRaw).trim() && targetFolderUidRaw !== '_default_';
          const validationErrors = [];
          if (!folderSelected) validationErrors.push('Shared folder selection is required');
          if (!recordTitle?.trim()) validationErrors.push('Title is required');
          if (validationErrors.length > 0) {
            const errorMessage = validationErrors.join('. ');
            log.debug(`Validation failed: ${errorMessage}`);
            return {
              statusCode: 200,
              type: 'application/vnd.microsoft.card.adaptive',
              value: await buildCreateSecretValidationCard(errorMessage),
            };
          }
        } else if (!recordTitle?.trim() || !recordLogin?.trim()) {
          const errors = [];
          if (!recordTitle?.trim()) errors.push('Title is required');
          if (!recordLogin?.trim()) errors.push('Login is required');
          const errorMessage = errors.join('. ');
          log.debug(`Validation failed: ${errorMessage}`);
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: buildApprovalValidationCard(errorMessage),
          };
        }

        if (isCreateSecretFlow) {
          log.debug('submit_create_record: create-secret (self-service) flow');
          const keeperClient = require('./services/keeperClient');
          const autoGenToggle = merged.autoGeneratePassword || 'false';
          const generatePassword = autoGenToggle === 'true' || autoGenToggle === true;
          const loginTrimmed = recordLogin?.trim() || '';

          const subfolderUidRaw = merged.targetSubfolderUid;
          const subfolderUid = subfolderUidRaw && String(subfolderUidRaw).trim() && subfolderUidRaw !== '_root_'
            ? String(subfolderUidRaw).trim()
            : null;
          const folderUid = subfolderUid
            || (targetFolderUidRaw && String(targetFolderUidRaw).trim() && targetFolderUidRaw !== '_default_'
              ? String(targetFolderUidRaw).trim()
              : null);

          const passwordValue = generatePassword ? '$GEN' : (recordPassword?.trim() || null);

          const result = await keeperClient.createRecord({
            title: recordTitle.trim(),
            ...(loginTrimmed ? { login: loginTrimmed } : {}),
            password: passwordValue,
            url: recordUrl?.trim() || null,
            notes: recordNotes?.trim() || null,
            generatePassword: generatePassword,
            selfDestructDuration: null,
            folderUid,
          });

          if (result.success) {
            // Fire-and-forget notification to approvers channel
            const channelService = getChannelService();
            if (channelService && channelService.isApprovalsChannelReady()) {
              const notificationCard = {
                type: 'AdaptiveCard',
                '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
                version: '1.5',
                body: [
                  { type: 'TextBlock', text: 'New Secret Created', weight: 'Bolder', size: 'Medium' },
                  {
                    type: 'FactSet',
                    facts: [
                      { title: 'Created by', value: merged.requesterName || 'Unknown' },
                      { title: 'Email', value: merged.requesterEmail || '—' },
                      { title: 'Title', value: recordTitle.trim() },
                      ...(result.recordUid ? [{ title: 'Record UID', value: result.recordUid }] : []),
                      ...(folderUid ? [{ title: 'Folder UID', value: folderUid }] : []),
                    ],
                  },
                ],
              };
              channelService.sendApprovalCardViaConnector(
                notificationCard,
                merged.approvalId || 'create-secret',
                `New secret created by **${merged.requesterName || 'Unknown'}**`
              ).catch((err) => log.error('Failed to send create-secret notification', err.message));
            }

            return {
              statusCode: 200,
              type: 'application/vnd.microsoft.card.adaptive',
              value: cards.buildCreateSecretSuccessCard({
                recordTitle: recordTitle.trim(),
                recordUid: result.recordUid,
              }),
            };
          }

          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: await buildCreateSecretValidationCard(result.error || 'Failed to create record'),
          };
        }

        // --- Approval-channel create flow (unchanged): async + proactive message to approvals ---
        const approvalPayload = merged;

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
            log.debug(`Background: Starting record creation${selfDestructEnabled ? ` (self-destruct: ${selfDestructDuration})` : ''}`);
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
            
            log.debug('Background: Record creation complete', result.success ? `UID: ${result.recordUid}` : `Error: ${result.error}`);
            
            // Build the result card with permission and duration dropdowns
            const { RECORD_PERMISSIONS, DURATION_OPTIONS } = require('./cards/constants');
            
            let resultCard;
            if (result.success) {
              // Build body elements
              const bodyElements = [
                { type: 'TextBlock', text: 'Record Created Successfully!', weight: 'Bolder', size: 'Large' },
                { type: 'TextBlock', text: `Requester: ${approvalPayload.requesterName || 'Unknown'}`, wrap: true },
                { type: 'TextBlock', text: `Justification: ${approvalPayload.justification || 'N/A'}`, wrap: true, isSubtle: true },
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
                version: '1.5',
                body: bodyElements,
                actions: [
                  {
                    type: 'Action.Execute',
                    title: 'Approve',
                    style: 'positive',
                    verb: 'approve_record',
                    data: { 
                      action: 'approve_record', 
                      approvalId: approvalPayload.approvalId || '', 
                      recordUid: result.recordUid, 
                      recordTitle: recordTitle.trim(), 
                      requesterId: approvalPayload.requesterId || '', 
                      requesterEmail: approvalPayload.requesterEmail || '', 
                      requesterName: approvalPayload.requesterName || '', 
                      justification: approvalPayload.justification || '',
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
                      approvalId: approvalPayload.approvalId || '', 
                      recordUid: result.recordUid, 
                      recordTitle: recordTitle.trim(), 
                      requesterId: approvalPayload.requesterId || '', 
                      requesterEmail: approvalPayload.requesterEmail || '', 
                      requesterName: approvalPayload.requesterName || '', 
                      justification: approvalPayload.justification || '',
                    },
                  },
                ],
              };
            } else {
              resultCard = {
                type: 'AdaptiveCard',
                '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
                version: '1.5',
                body: [
                  { type: 'TextBlock', text: 'Error Creating Record', weight: 'Bolder', size: 'Large', color: 'Attention' },
                  { type: 'TextBlock', text: result.error || 'Unknown error occurred', wrap: true },
                ],
                actions: [
                  { type: 'Action.Execute', title: 'Try Again', verb: 'show_create_form', data: approvalPayload },
                ],
              };
            }
            
            if (channelService && channelService.isApprovalsChannelReady()) {
              try {
                const sent = await channelService.sendApprovalCardViaConnector(
                  resultCard, 
                  approvalPayload.approvalId || 'create-record',
                  result.success ? `Record "${recordTitle.trim()}" created successfully!` : null
                );
                log.debug('Background: Sent result card via channelService', sent.success);
              } catch (sendError) {
                log.error('Background: Error sending via channelService', sendError.message);
              }
            } else {
              log.error('Background: ChannelService not available for proactive messaging');
            }
          } catch (bgError) {
            log.error('Background: Error in record creation', bgError.message);
          }
        })();
        
        log.debug('Returning processing card immediately');
        return {
          statusCode: 200,
          type: 'application/vnd.microsoft.card.adaptive',
          value: {
            type: 'AdaptiveCard',
            '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
            version: '1.5',
            body: [
              { type: 'TextBlock', text: 'Creating Record...', weight: 'Bolder', size: 'Large' },
              { type: 'TextBlock', text: `Title: ${recordTitle.trim()}`, wrap: true },
              { type: 'TextBlock', text: 'Please wait. The result will appear below shortly.', wrap: true, isSubtle: true },
            ],
            actions: [],
          },
        };
      } catch (error) {
        log.error('Error handling submit_create_record', { message: error.message, stack: error.stack });
        return { statusCode: 500, body: error.message };
      }
    }
    
    if (verb === 'refresh_shared_folders') {
      try {
        log.debug('Processing refresh_shared_folders');
        const keeperClient = require('./services/keeperClient');
        const cards = require('./cards');

        const v = activity.value || {};
        const merged = { ...data, ...(v.action?.data || {}), ...(v.data || {}) };

        const folderResult = await keeperClient.getSharedFolderChoicesForEmail(merged.requesterEmail || '');

        const refreshedCard = cards.buildRecordCreationCard({
          approvalId: merged.approvalId,
          requesterName: merged.requesterName,
          requesterId: merged.requesterId,
          requesterEmail: merged.requesterEmail,
          requesterAadObjectId: merged.requesterAadObjectId,
          justification: merged.justification,
          identifier: merged.identifier,
          originalRecordTitle: merged.originalRecordTitle,
          searchQuery: merged.searchQuery,
          createSecretFlow: true,
          sharedFolderChoices: folderResult.choiceSetChoices,
          sharedFoldersLoadError: folderResult.error,
          noSharedFoldersForUser: !!folderResult.noSharedFoldersForUser,
          selectedTargetFolderUid: '_default_',
          recordTitle: merged.recordTitle || '',
          recordLogin: merged.recordLogin || '',
          recordPassword: merged.recordPassword || '',
          recordUrl: merged.recordUrl || '',
          recordNotes: merged.recordNotes || '',
        });

        return {
          statusCode: 200,
          type: 'application/vnd.microsoft.card.adaptive',
          value: refreshedCard,
        };
      } catch (error) {
        log.error('Error handling refresh_shared_folders', error);
        return { statusCode: 500, body: error.message };
      }
    }

    if (verb === 'load_subfolders') {
      try {
        log.debug('Processing load_subfolders');
        const keeperClient = require('./services/keeperClient');
        const cards = require('./cards');

        const v = activity.value || {};
        const merged = { ...data, ...(v.action?.data || {}), ...(v.data || {}) };
        const selectedFolder = v.targetFolderUid || merged.targetFolderUid;

        if (!selectedFolder || !String(selectedFolder).trim()) {
          log.debug('load_subfolders: no folder selected');
          return { statusCode: 200 };
        }

        const folderResult = await keeperClient.getSharedFolderChoicesForEmail(merged.requesterEmail || '');
        const subfolderResult = await keeperClient.getSubfoldersForSharedFolder(String(selectedFolder).trim());

        const refreshedCard = cards.buildRecordCreationCard({
          approvalId: merged.approvalId,
          requesterName: merged.requesterName,
          requesterId: merged.requesterId,
          requesterEmail: merged.requesterEmail,
          requesterAadObjectId: merged.requesterAadObjectId,
          justification: merged.justification,
          identifier: merged.identifier,
          originalRecordTitle: merged.originalRecordTitle,
          searchQuery: merged.searchQuery,
          createSecretFlow: true,
          sharedFolderChoices: folderResult.choiceSetChoices,
          sharedFoldersLoadError: folderResult.error,
          noSharedFoldersForUser: !!folderResult.noSharedFoldersForUser,
          selectedTargetFolderUid: selectedFolder,
          subfolderChoices: subfolderResult.choices,
          subfolderLoadError: subfolderResult.error,
          selectedSubfolderUid: v.targetSubfolderUid || merged.targetSubfolderUid || '',
          recordTitle: v.recordTitle || merged.recordTitle || '',
          recordLogin: v.recordLogin || merged.recordLogin || '',
          recordPassword: v.recordPassword || merged.recordPassword || '',
          recordUrl: v.recordUrl || merged.recordUrl || '',
          recordNotes: v.recordNotes || merged.recordNotes || '',
        });

        return {
          statusCode: 200,
          type: 'application/vnd.microsoft.card.adaptive',
          value: refreshedCard,
        };
      } catch (error) {
        log.error('Error handling load_subfolders', error);
        return { statusCode: 500, body: error.message };
      }
    }

    if (verb === 'cancel_create_secret') {
      log.debug('Processing cancel_create_secret');
      return {
        statusCode: 200,
        type: 'application/vnd.microsoft.card.adaptive',
        value: {
          type: 'AdaptiveCard',
          '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
          version: '1.5',
          body: [
            { type: 'TextBlock', text: 'Cancelled', weight: 'Bolder', size: 'Medium' },
            { type: 'TextBlock', text: 'No record was created.', wrap: true, isSubtle: true },
          ],
        },
      };
    }

    if (verb === 'cancel_create_form') {
      try {
        log.debug('Processing cancel_create_form');
        const { handleCancelCreateForm } = require('./handlers/approvalHandler');
        const searchCard = await handleCancelCreateForm(data);
        
        return {
          statusCode: 200,
          type: 'application/vnd.microsoft.card.adaptive',
          value: searchCard,
        };
      } catch (error) {
        log.error('Error handling cancel_create_form', error);
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
          log.error('No record selected');
          return { statusCode: 400, body: 'No record selected' };
        }
        
        let selectedRecord;
        try {
          selectedRecord = JSON.parse(selectedRecordJson);
        } catch (e) {
          log.error('Failed to parse selected record', e);
          return { statusCode: 400, body: 'Invalid record selection' };
        }
        
        log.debug('Approving selected record', selectedRecord);
        
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
        
        if (result && result.error) {
          log.error('approve_selected_record returned error', result.error);
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: {
              type: 'AdaptiveCard',
              '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
              version: '1.5',
              body: [
                { type: 'TextBlock', text: 'Approval Failed', weight: 'Bolder', size: 'Large', color: 'Attention' },
                { type: 'TextBlock', text: result.error || 'An error occurred.', wrap: true },
              ],
            },
          };
        }
        
        return { statusCode: 200 };
      } catch (error) {
        log.error('Error handling approve_selected_record', error);
        return { statusCode: 500, body: error.message };
      }
    }
    
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
          log.error('No folder selected');
          return { statusCode: 400, body: 'No folder selected' };
        }
        
        let selectedFolder;
        try {
          selectedFolder = JSON.parse(selectedFolderJson);
        } catch (e) {
          log.error('Failed to parse selected folder', e);
          return { statusCode: 400, body: 'Invalid folder selection' };
        }
        
        log.debug('Approving selected folder', selectedFolder);
        
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
        
        if (result && result.error) {
          log.error('approve_selected_folder returned error', result.error);
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: {
              type: 'AdaptiveCard',
              '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
              version: '1.5',
              body: [
                { type: 'TextBlock', text: 'Approval Failed', weight: 'Bolder', size: 'Large', color: 'Attention' },
                { type: 'TextBlock', text: result.error || 'An error occurred.', wrap: true },
              ],
            },
          };
        }
        
        return { statusCode: 200 };
      } catch (error) {
        log.error('Error handling approve_selected_folder', error);
        return { statusCode: 500, body: error.message };
      }
    }
    
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
          log.error('No record selected for share');
          return { statusCode: 400, body: 'No record selected' };
        }
        
        let selectedRecord;
        try {
          selectedRecord = JSON.parse(selectedRecordJson);
        } catch (e) {
          log.error('Failed to parse selected record', e);
          return { statusCode: 400, body: 'Invalid record selection' };
        }
        
        log.debug('Creating one-time share for selected record', selectedRecord);
        
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
        
        if (result && result.error) {
          log.error('approve_selected_share returned error', result.error);
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: {
              type: 'AdaptiveCard',
              '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
              version: '1.5',
              body: [
                { type: 'TextBlock', text: 'Share Creation Failed', weight: 'Bolder', size: 'Large', color: 'Attention' },
                { type: 'TextBlock', text: result.error || 'An error occurred.', wrap: true },
              ],
            },
          };
        }
        
        return { statusCode: 200 };
      } catch (error) {
        log.error('Error handling approve_selected_share', error);
        return { statusCode: 500, body: error.message };
      }
    }
    
    if (action?.startsWith('approve_') || action?.startsWith('deny_')) {
      try {
        log.debug(`Processing ${action} from adaptiveCard/action (Universal Action)`);
        
        if (action.includes('pedm')) {
          const resultCard = await handlers.routePedmAction(context, data);
          if (resultCard) {
            log.debug(`EPM action ${action} completed, updating card in-place`);
            return {
              statusCode: 200,
              type: 'application/vnd.microsoft.card.adaptive',
              value: resultCard,
            };
          }
          return { statusCode: 200 };
        }
        
        if (action.includes('device')) {
          const resultCard = await handlers.routeDeviceAction(context, data);
          if (resultCard) {
            log.debug(`Device action ${action} completed, updating card in-place`);
            return {
              statusCode: 200,
              type: 'application/vnd.microsoft.card.adaptive',
              value: resultCard,
            };
          }
          return { statusCode: 200 };
        }
        
        const result = await handlers.routeApprovalActionWithCardResponse(context, data);
        
        log.debug(`Action ${action} completed`, { hasCard: !!result?.updatedCard, hasError: !!result?.error });
        
        if (result && result.updatedCard) {
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: result.updatedCard,
          };
        }
        
        if (result && result.error) {
          log.error(`Action ${action} returned error`, result.error);
          const errorCard = {
            type: 'AdaptiveCard',
            '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
            version: '1.5',
            body: [
              { type: 'TextBlock', text: 'Action Failed', weight: 'Bolder', size: 'Large', color: 'Attention' },
              { type: 'TextBlock', text: result.error || 'An error occurred while processing the request.', wrap: true },
            ],
            actions: result.keepCardActive ? [] : [],
          };
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: errorCard,
          };
        }
        
        return { statusCode: 200 };
      } catch (error) {
        log.error(`Error handling ${action}`, error);
        return { statusCode: 500, body: error.message };
      }
    }
  }
  
  log.debug('Unhandled invoke name', invokeName);
  
  // Return undefined to let other handlers process
  return undefined;
});

// ==================== Task Module Handler ====================

app.on("task/fetch", async (context) => {
  const activity = context.activity;
  
  log.debug('Task module fetch (legacy)', activity.value);
  
  try {
    const { handleTaskFetch } = require('./handlers/taskModuleHandler');
    const response = await handleTaskFetch(context, activity);
    return response;
  } catch (error) {
    log.error('Error handling task/fetch', error);
    return {
      task: {
        type: 'message',
        value: 'Error opening task module: ' + error.message,
      },
    };
  }
});

app.on("task/submit", async (context) => {
  const activity = context.activity;
  
  log.debug('Task module submit (legacy)', activity.value);
  
  try {
    const { handleTaskSubmit } = require('./handlers/taskModuleHandler');
    const response = await handleTaskSubmit(context, activity);
    return response;
  } catch (error) {
    log.error('Error handling task/submit', error);
    return {
      task: {
        type: 'message',
        value: 'Error processing task module: ' + error.message,
      },
    };
  }
});

app.on("cardAction", async (context) => {
  const activity = context.activity;
  const data = activity.value || {};
  
  log.debug('Card action received', data);

  const action = data.action;

  if (!action) {
    log.debug('No action in card data, skipping');
    return;
  }

  if (data.msteams && data.msteams.type === 'task/fetch') {
    log.debug('Task module request detected in cardAction, should be handled by invoke handler');
    return undefined;
  }

  try {
    if (action?.startsWith('approve_') || action?.startsWith('deny_')) {
      log.debug(`Routing approval/denial action: ${action}`);
      if (action.includes('record') || action.includes('folder') || action.includes('share')) {
        await handlers.routeApprovalAction(context, data);
        log.debug(`Action ${action} completed`);
      } else if (action.includes('pedm')) {
        const resultCard = await handlers.routePedmAction(context, data);
        if (resultCard) {
          log.debug(`EPM action ${action} completed, updating card in-place`);
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: resultCard,
          };
        }
      } else if (action.includes('device')) {
        const resultCard = await handlers.routeDeviceAction(context, data);
        if (resultCard) {
          log.debug(`Device action ${action} completed, updating card in-place`);
          return {
            statusCode: 200,
            type: 'application/vnd.microsoft.card.adaptive',
            value: resultCard,
          };
        }
      }
    } else {
      log.debug(`Unknown action: ${action}`);
    }
  } catch (error) {
    log.error('Error handling card action', error);
    await context.send(`Error processing action: ${error.message}`);
  }
});

// ==================== Background Pollers ====================

// Initialize pollers (they will only start if enabled in config)
let epmPoller = null;
let devicePoller = null;

const startPollers = () => {
  const currentConfig = getConfig();
  if (currentConfig.pedm.enabled) {
    epmPoller = new EpmPoller(app);
    epmPoller.start();
    log.info('EPM poller initialized');
  }

  if (currentConfig.deviceApproval.enabled) {
    devicePoller = new DevicePoller(app);
    devicePoller.start();
    log.info('Device poller initialized');
  }
};

// Stop pollers on shutdown
const stopPollers = () => {
  if (epmPoller) epmPoller.stop();
  if (devicePoller) devicePoller.stop();
};

process.on('SIGINT', () => {
  log.info('Shutting down...');
  stopPollers();
  process.exit(0);
});

process.on('SIGTERM', () => {
  log.info('Shutting down...');
  stopPollers();
  process.exit(0);
});

// ==================== App Ready Hook ====================

const originalStart = app.start.bind(app);
app.start = async (...args) => {
  const result = await originalStart(...args);
  
  log.info('='.repeat(60));
  log.info('Starting Keeper Commander Teams Bot');
  log.info('='.repeat(60));
  
  const keeperClient = require('./services/keeperClient');
  const currentConfig = getConfig();
  const serviceUrl = currentConfig.keeper?.serviceUrl || 'http://localhost:7001/api/v2/';
  
  try {
    const isHealthy = await keeperClient.healthCheck();
    
    if (isHealthy) {
      log.info('Keeper Service Mode is accessible');
    } else {
      log.warn(`Cannot reach Keeper Service Mode at ${serviceUrl}. The bot will start but commands may fail.`);
    }
  } catch (error) {
    log.warn(`Cannot reach Keeper Service Mode at ${serviceUrl}. Error: ${error.message}. The bot will start but commands may fail.`);
  }
  
  log.info('='.repeat(60));
  
  setTimeout(() => {
    startPollers();
  }, 2000);
  
  return result;
};

module.exports = app;
