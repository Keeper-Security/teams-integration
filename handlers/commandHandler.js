/**
 * Command Handler for Keeper Security Bot
 * 
 * Handles commands from users:
 * - keeper-request-record <name> <justification>
 * - keeper-request-folder <name> <justification>
 * - keeper-one-time-share <name> [justification]
 * - keeper-create-secret "<title>" [notes]  (title required; notes optional, pre-fills the form)
 * - help
 */

const keeperClient = require('../services/keeperClient');
const { getChannelService, createLogger } = require('../services');
const graphService = require('../services/graphService');
const cards = require('../cards');
const config = require('../config');
const { isUid, looksLikeInvalidUid, isPamRecordError } = require('../utils/helpers');

const log = createLogger('CommandHandler');

/**
 * Generate a unique approval ID
 * @returns {string} - Short unique ID
 */
function generateApprovalId() {
  return Math.random().toString(36).substring(2, 10);
}

/**
 * Parse command text into command name and arguments
 * @param {string} text - Raw message text
 * @returns {Object} - { command, args }
 */
function parseCommand(text) {
  const trimmed = text.trim();
  
  // Handle commands with or without leading slash
  const normalized = trimmed.startsWith('/') ? trimmed.substring(1) : trimmed;
  
  // Split into command and remaining args
  const [command, ...rest] = normalized.split(/\s+/);
  const argsText = rest.join(' ');
  
  return {
    command: command.toLowerCase(),
    argsText,
    args: rest,
  };
}

/**
 * Parse uid and justification from args text
 * Handles quoted strings for names with spaces
 * @param {string} argsText - Raw arguments text
 * @returns {Object} - { uid, justification }
 */
function parseUidAndJustification(argsText) {
  const trimmed = argsText.trim();
  
  if (!trimmed) {
    return { uid: null, justification: '' };
  }
  
  // Check if starts with quote (name with spaces)
  // Support both straight quotes and smart/curly quotes
  // Smart quotes: " (\u201C), " (\u201D), ' (\u2018), ' (\u2019)
  const quoteChars = ['"', "'", '\u201C', '\u201D', '\u2018', '\u2019'];
  const firstChar = trimmed[0];
  
  if (quoteChars.includes(firstChar)) {
    // Map smart quotes to their closing pairs
    const closingQuote = firstChar === '\u201C' ? '\u201D' : 
                         firstChar === '\u2018' ? '\u2019' : firstChar;
    const endQuote = trimmed.indexOf(closingQuote, 1);
    if (endQuote > 0) {
      const uid = trimmed.substring(1, endQuote).trim();
      const justification = trimmed.substring(endQuote + 1).trim();
      // Return null if uid is empty after trimming
      return { uid: uid || null, justification };
    }
  }
  
  // Otherwise, first word is uid, rest is justification
  const [uid, ...rest] = trimmed.split(/\s+/);
  return { uid: uid || null, justification: rest.join(' ') };
}

/**
 * Get user info from activity
 * Fetches real email address from Microsoft Graph API
 * @param {Object} activity - Teams activity
 * @returns {Promise<Object>} - { userId, userName, userEmail }
 */
async function getUserInfo(activity) {
  const from = activity.from || {};
  const userInfo = {
    teamsUserId: from.id || 'unknown',
    userName: from.name || 'Unknown User',
    userEmail: null, // Will be fetched from Graph API
  };

  // Try to get email from Microsoft Graph API
  const aadObjectId = from.aadObjectId;
  if (aadObjectId) {
    try {
      const email = await graphService.getUserEmail(aadObjectId);
      if (email) {
        userInfo.userEmail = email;
        log.debug(`Fetched email for user ${from.name}: ${email}`);
      } else {
        log.warn(`No email found for user ${from.name} (AAD ID: ${aadObjectId})`);
      }
    } catch (error) {
      log.error(`Error fetching email for user ${from.name}`, error.message);
      userInfo.userEmail = from.email || from.userPrincipalName || null;
    }
  } else {
    userInfo.userEmail = from.email || from.userPrincipalName || null;
    if (!userInfo.userEmail) {
      log.warn(`No AAD Object ID found for user ${from.name}, cannot fetch email`);
    }
  }

  return userInfo;
}

/**
 * Handle the request-record command
 * @param {Object} context - Teams turn context
 * @param {string} argsText - Command arguments (uid/name + justification)
 * @returns {Promise<void>}
 */
async function handleRequestRecord(context, argsText) {
  const { uid, justification } = parseUidAndJustification(argsText);
  
  if (!uid || !uid.trim()) {
    await context.send('**Usage:** `keeper-request-record <record-name> <justification>`\n\nExample: `keeper-request-record AWS-Prod Need for deployment`');
    return;
  }
  
  if (!justification || !justification.trim()) {
    await context.send('Please provide a justification for your access request.\n\n**Usage:** `keeper-request-record <record-name> <justification>`');
    return;
  }
  
  const trimmedUid = uid.trim();
  
  // Check if it looks like an invalid UID (wrong length but UID-like pattern)
  if (looksLikeInvalidUid(trimmedUid)) {
    await context.send(`**Invalid UID Format**\n\nThe identifier \`${trimmedUid}\` looks like a UID but has an invalid length (${trimmedUid.length} characters).\n\nKeeper UIDs are typically 22 characters. Please verify and try again.`);
    return;
  }
  
  // Check if identifier is a UID or description
  const isUidFormat = isUid(trimmedUid);
  let record = null;
  let recordUid = null;
  let recordTitle = uid; // Default to the identifier
  
  if (isUidFormat) {
    // Try to get record by UID
    record = await keeperClient.getRecordByUid(trimmedUid);
    if (record) {
      recordUid = record.uid;
      recordTitle = record.title;
      
      // Check if it's actually a folder, not a record
      const recordType = (record.recordType || record.record_type || '').toLowerCase();
      if (recordType.includes('folder') || recordType === 'shared_folder' || recordType === 'user_folder') {
        await context.send(`**Invalid UID Type**\n\nThe UID \`${trimmedUid}\` is a **folder**, not a record.\n\nPlease use \`keeper-request-folder ${trimmedUid} ${justification}\` instead.`);
        return;
      }
    }
  }
  
  // If not found by UID or not a UID format, treat as description
  if (!record && isUidFormat) {
    // UID format but not found - might be invalid
    await context.send(`**Record Not Found**\n\nNo record found with UID: \`${trimmedUid}\`\n\nPlease verify the UID and try again.`);
    return;
  }
  
  // Get user info
  const userInfo = await getUserInfo(context.activity);
  const approvalId = generateApprovalId();
  
  // Build approval card
  // If not a UID or UID not found, pass isUid=false and identifier
  const approvalCard = cards.createRecordApprovalCard({
    approvalId,
    recordUid: recordUid || uid, // Use actual UID if found, otherwise identifier
    recordTitle: recordTitle,
    recordType: record?.recordType || 'Unknown',
    requesterName: userInfo.userName,
    requesterId: userInfo.teamsUserId,
    requesterEmail: userInfo.userEmail,
    requesterAadObjectId: context.activity.from?.aadObjectId, // Extract AAD Object ID for email fetching fallback
    justification,
    isUid: isUidFormat && !!record, // True only if UID format AND found
    identifier: uid, // Always pass the original identifier
    isNsf: !!record?.isNsf,
  });
  
  // Try to send to approvals channel
  const channelService = getChannelService();
  let sentToChannel = false;
  
  if (channelService) {
    const status = channelService.getStatus();
    const canSend = status.approvalsChannelReady && status.appReady;
    
    if (canSend) {
      // Use connector-based method for sending - this allows us to update the card later
      const result = await channelService.sendApprovalCardViaConnector(
        approvalCard,
        approvalId,
        `New record access request from **${userInfo.userName}**`
      );
      sentToChannel = result.success;
      
      if (result.activityId) {
        log.debug(`Approval card sent with activityId: ${result.activityId}`);
      }
    } else {
      log.debug('Cannot send to approvals channel', { approvalsReady: status.approvalsChannelReady, appReady: status.appReady });
    }
  }
  
  if (sentToChannel) {
    await context.send(
      '**Record access request submitted!**  \n' +
      `• **Request ID:** \`${approvalId}\`  \n` +
      `• **Record:** \`${uid}\`  \n` +
      `• **Justification:** ${justification}  \n` +
      'Your request has been sent to the **approvals channel** for review.  \n' +
      'You will be notified when an administrator approves or denies your request.'
    );
  } else {
    // Fallback: send to current conversation (for testing or when channel not configured)
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: approvalCard,
      }],
    });
    
    const displayTitle = record?.title || recordTitle || uid;
    await context.send(
      'Access request submitted for **' + displayTitle + '**\n\n' +
      'An administrator will review your request.\n\n' +
      '_Note: Configure APPROVALS_CHANNEL_ID to route requests to a dedicated channel._'
    );
  }
}

/**
 * Handle the request-folder command
 * @param {Object} context - Teams turn context
 * @param {string} argsText - Command arguments (uid/name + justification)
 * @returns {Promise<void>}
 */
async function handleRequestFolder(context, argsText) {
  const { uid, justification } = parseUidAndJustification(argsText);
  
  if (!uid || !uid.trim()) {
    await context.send('**Usage:** `keeper-request-folder <folder-name> <justification>`\n\nExample: `keeper-request-folder "Engineering Creds" Project onboarding`');
    return;
  }
  
  if (!justification || !justification.trim()) {
    await context.send('Please provide a justification for your access request.\n\n**Usage:** `keeper-request-folder <folder-name> <justification>`');
    return;
  }
  
  const trimmedUid = uid.trim();
  
  // Check if it looks like an invalid UID (wrong length but UID-like pattern)
  if (looksLikeInvalidUid(trimmedUid)) {
    await context.send(`**Invalid UID Format**\n\nThe identifier \`${trimmedUid}\` looks like a UID but has an invalid length (${trimmedUid.length} characters).\n\nKeeper UIDs are typically 22 characters. Please verify and try again.`);
    return;
  }
  
  // Check if identifier is a UID or description
  const isUidFormat = isUid(trimmedUid);
  let folder = null;
  let folderUid = null;
  let folderName = uid; // Default to the identifier
  
  if (isUidFormat) {
    // Try to get folder by UID
    folder = await keeperClient.getFolderByUid(trimmedUid);
    if (folder) {
      folderUid = folder.uid;
      folderName = folder.name;
      
      // Check if it's actually a record, not a folder
      const folderType = (folder.folderType || folder.folder_type || '').toLowerCase();
      if (folderType === 'record' || folderType === 'login' || folderType === 'general') {
        await context.send(`**Invalid UID Type**\n\nThe UID \`${trimmedUid}\` is a **record**, not a folder.\n\nPlease use \`keeper-request-record ${trimmedUid} ${justification}\` instead.`);
        return;
      }
    }
  }
  
  // If not found by UID or not a UID format, treat as description
  if (!folder && isUidFormat) {
    // UID format but not found - might be invalid
    await context.send(`**Folder Not Found**\n\nNo folder found with UID: \`${trimmedUid}\`\n\nPlease verify the UID and try again.`);
    return;
  }
  
  // Get user info
  const userInfo = await getUserInfo(context.activity);
  const approvalId = generateApprovalId();
  
  // Build approval card
  // If not a UID or UID not found, pass isUid=false and identifier
  const approvalCard = cards.createFolderApprovalCard({
    approvalId,
    folderUid: folderUid || uid, // Use actual UID if found, otherwise identifier
    folderName: folderName,
    folderType: folder?.folderType || 'shared_folder',
    requesterName: userInfo.userName,
    requesterId: userInfo.teamsUserId,
    requesterEmail: userInfo.userEmail,
    requesterAadObjectId: context.activity.from?.aadObjectId, // Extract AAD Object ID for email fetching fallback
    justification,
    isUid: isUidFormat && !!folder, // True only if UID format AND found
    identifier: uid, // Always pass the original identifier
    isNsf: !!folder?.isNsf,
  });
  
  // Try to send to approvals channel
  const channelService = getChannelService();
  let sentToChannel = false;
  
  if (channelService) {
    const status = channelService.getStatus();
    const canSend = status.approvalsChannelReady && status.appReady;
    
    if (canSend) {
      // Use connector-based method for sending - this allows us to update the card later
      const result = await channelService.sendApprovalCardViaConnector(
        approvalCard,
        approvalId,
        `New folder access request from **${userInfo.userName}**`
      );
      sentToChannel = result.success;
      
      if (result.activityId) {
        log.debug(`Folder approval card sent with activityId: ${result.activityId}`);
      }
    } else {
      log.debug('Cannot send to approvals channel', { approvalsReady: status.approvalsChannelReady, appReady: status.appReady });
    }
  }

  if (sentToChannel) {
    await context.send(
      '**Folder access request submitted!**  \n' +
      `• **Request ID:** \`${approvalId}\`  \n` +
      `• **Folder:** \`${uid}\`  \n` +
      `• **Justification:** ${justification}  \n` +
      'Your request has been sent to the **approvals channel** for review.  \n' +
      'You will be notified when an administrator approves or denies your request.'
    );
  } else {
    // Fallback: send to current conversation (for testing or when channel not configured)
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: approvalCard,
      }],
    });
    
    const displayName = folder?.name || folderName || uid;
    await context.send(
      'Access request submitted for folder **' + displayName + '**\n\n' +
      'An administrator will review your request.\n\n' +
      '_Note: Configure APPROVALS_CHANNEL_ID to route requests to a dedicated channel._'
    );
  }
}

/**
 * Handle the share (one-time-share) command
 * @param {Object} context - Teams turn context
 * @param {string} argsText - Command arguments (uid/name + optional justification)
 * @returns {Promise<void>}
 */
async function handleShare(context, argsText) {
  const { uid, justification } = parseUidAndJustification(argsText);
  
  if (!uid || !uid.trim()) {
    await context.send('**Usage:** `keeper-one-time-share <record-name> [justification]`\n\nExample: `keeper-one-time-share AWS-Prod Share with contractor`');
    return;
  }
  
  if (!justification || !justification.trim()) {
    await context.send('**Justification is required.**\n\n**Usage:** `keeper-one-time-share "<record-name>" <justification>`\n\nExample: `keeper-one-time-share "AWS-Prod" Share with contractor for audit`');
    return;
  }
  
  const trimmedUid = uid.trim();
  
  // Check if it looks like an invalid UID (wrong length but UID-like pattern)
  if (looksLikeInvalidUid(trimmedUid)) {
    await context.send(`**Invalid UID Format**\n\nThe identifier \`${trimmedUid}\` looks like a UID but has an invalid length (${trimmedUid.length} characters).\n\nKeeper UIDs are typically 22 characters. Please verify and try again.`);
    return;
  }
  
  // Check if input is a valid UID (22 char base64)
  const inputIsUid = isUid(trimmedUid);
  const requiresApproval = config.features.requireShareApproval;
  
  // For approval flow with description, we don't need to fetch record details upfront
  // The admin will search and select the correct record
  let record = null;
  let recordUid = null;
  let recordTitle = uid;
  
  // Only fetch record details if it's a UID (for validation and display)
  if (inputIsUid) {
    record = await keeperClient.getRecordByUid(trimmedUid);
    if (record) {
      recordUid = record.uid;
      recordTitle = record.title;
      
      // Check if it's a folder (one-time-share only works for records)
      const recordType = (record.recordType || record.record_type || '').toLowerCase();
      if (recordType.includes('folder') || recordType === 'shared_folder' || recordType === 'user_folder') {
        await context.send(`**Invalid UID Type**\n\nThe UID \`${trimmedUid}\` is a **folder**, not a record.\n\nOne-time share is only available for records. Please use \`keeper-request-folder\` for folder access.`);
        return;
      }
      
      // Check if it's a PAM record (one-time-share not available for PAM records)
      if (recordType.includes('pam') || recordType.startsWith('pam')) {
        await context.send(`**One-Time Share Not Available**\n\nThe record \`${recordTitle}\` is a PAM record.\n\nOne-Time Shares are currently not available for PAM records. Please use \`keeper-request-record\` to request direct access instead.`);
        return;
      }
    } else {
      await context.send(`**Record Not Found**\n\nNo record found with UID: \`${trimmedUid}\`\n\nPlease verify the UID and try again.`);
      return;
    }
  }
  
  // Check if share approval is required
  if (requiresApproval) {
    const userInfo = await getUserInfo(context.activity);
    const approvalId = generateApprovalId();
    
    // Build approval card - similar to record request flow
    // If UID: direct approval with details
    // If description: approval with search button
    const approvalCard = cards.buildOneTimeShareApprovalCard({
      approvalId,
      recordUid: recordUid,
      recordTitle: recordTitle,
      requesterName: userInfo.userName,
      requesterId: userInfo.teamsUserId,
      requesterEmail: userInfo.userEmail,
      requesterAadObjectId: context.activity.from?.aadObjectId,
      justification: justification,
      isUid: inputIsUid,
      identifier: uid,
    });
    
    // Try to send to approvals channel with activity ID tracking for in-place updates
    const channelService = getChannelService();
    let sentToChannel = false;
    
    if (channelService) {
      const status = channelService.getStatus();
      const canSend = status.approvalsChannelReady && status.appReady;
      
      if (canSend) {
        // Use sendApprovalCardViaConnector to store activity ID for in-place updates
        const sendResult = await channelService.sendApprovalCardViaConnector(
          approvalCard,
          approvalId,
          `New one-time share request from **${userInfo.userName}**`
        );
        sentToChannel = sendResult.success;
        
        if (sendResult.activityId) {
          log.debug(`Share approval card sent with activityId: ${sendResult.activityId}`);
        }
      } else {
        log.debug('Cannot send to approvals channel', { approvalsReady: status.approvalsChannelReady, appReady: status.appReady });
      }
    }
    
    if (sentToChannel) {
      // For description-based requests, only show the identifier (not resolved title)
      await context.send(
        '**One-Time Share request submitted!**  \n' +
        `• **Request ID:** \`${approvalId}\`  \n` +
        `• **Record:** \`${uid}\`  \n` +
        `• **Justification:** ${justification}  \n` +
        'Your request has been sent to the **approvals channel** for review.  \n' +
        'Once approved, you will receive the share link via DM.'
      );
    } else {
      await context.send({
        type: 'message',
        attachments: [{
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: approvalCard,
        }],
      });
      
      await context.send(
        'Share request submitted for **' + (record?.title || uid) + '**\n\n' +
        'An administrator will review your request.\n\n' +
        '_Note: Configure APPROVALS_CHANNEL_ID to route requests to a dedicated channel._'
      );
    }
    return;
  }
  
  // Create the share directly (no approval required)
  await context.send('Creating one-time share link...');
  
  const result = await keeperClient.createOneTimeShare(record?.uid || uid, 86400, false);
  
  if (result.success) {
    const resultCard = cards.createShareResultCard({
      recordTitle: record?.title || uid,
      shareUrl: result.shareUrl,
      expiresAt: result.expiresAt,
    });
    
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: resultCard,
      }],
    });
  } else {
    await context.send('Failed to create share link: ' + result.error);
  }
}

/**
 * Handle the search command
 * @param {Object} context - Teams turn context
 * @param {string} argsText - Search query
 * @param {string} searchType - 'record' or 'folder'
 * @returns {Promise<void>}
 */
async function handleSearch(context, argsText, searchType = 'record') {
  const query = argsText.trim();
  
  if (!query) {
    const typeLabel = searchType === 'record' ? 'records' : 'folders';
    await context.send('**Usage:** `search-' + typeLabel + ' <query>`\n\nExample: `search-' + typeLabel + ' production`');
    return;
  }
  
  if (!config.features.enableSearch) {
    await context.send('Search is not enabled. Contact your administrator.');
    return;
  }
  
  await context.send('Searching...');
  
  let results;
  if (searchType === 'record') {
    results = await keeperClient.searchRecords(query, 10);
  } else {
    results = await keeperClient.searchFolders(query, 10);
  }
  
  if (results.length === 0) {
    await context.send('No ' + searchType + 's found matching "' + query + '"');
    return;
  }
  
  // Build results card
  const resultsCard = cards.createSearchResultsCard({
    searchType,
    query,
    results,
  });
  
  await context.send({
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: resultsCard,
    }],
  });
}

/**
 * Handle the help command
 * @param {Object} context - Teams turn context
 * @returns {Promise<void>}
 */
async function handleHelp(context) {
  const helpCard = cards.createHelpCard();
  
  await context.send({
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: helpCard,
    }],
  });
}

/**
 * Handle the status command
 * @param {Object} context - Teams turn context
 * @returns {Promise<void>}
 */
async function handleStatus(context) {
  await context.send('Checking Keeper connection...');
  
  const isHealthy = await keeperClient.healthCheck();
  
  if (isHealthy) {
    await context.send('**Keeper Status:** Connected\n\n' +
      '• Service URL: `' + (config.keeper?.serviceUrl || 'Not configured') + '`\n' +
      '• EPM Polling: ' + (config.pedm?.enabled ? 'Enabled' : 'Disabled') + '\n' +
      '• Device Approval: ' + (config.deviceApproval?.enabled ? 'Enabled' : 'Disabled') + '\n' +
      '• Search: ' + (config.features?.enableSearch ? 'Enabled' : 'Disabled')
    );
  } else {
    await context.send('**Keeper Status:** Disconnected\n\n' +
      'Unable to connect to Keeper Commander Service Mode.\n' +
      'Service URL: `' + (config.keeper?.serviceUrl || 'Not configured') + '`\n\n' +
      'Please check that Keeper Commander is running in Service Mode.'
    );
  }
}

/**
 * Show the create-secret Adaptive Card (title, login, password, URL, notes; optional shared-folder target; no self-destruct).
 * Optional message args pre-fill title and notes.
 *
 * @param {Object} context - Teams turn context
 * @param {string} argsText - Optional: quoted title + optional notes for pre-fill
 * @returns {Promise<void>}
 */
async function handleCreateSecret(context, argsText) {
  const { uid: prefillTitle, justification: prefillNotes } = parseUidAndJustification(argsText || '');
  const userInfo = await getUserInfo(context.activity);
  const approvalId = generateApprovalId();

  const {
    choiceSetChoices: sharedFolderChoices,
    error: sharedFoldersLoadError,
    noSharedFoldersForUser,
  } = await keeperClient.getSharedFolderChoicesForEmail(userInfo.userEmail);

  const formCard = cards.buildRecordCreationCard({
    approvalId,
    requesterName: userInfo.userName,
    requesterId: userInfo.teamsUserId,
    requesterEmail: userInfo.userEmail,
    requesterAadObjectId: context.activity.from?.aadObjectId,
    justification: '',
    identifier: '',
    originalRecordTitle: '',
    searchQuery: prefillTitle?.trim() || '',
    createSecretFlow: true,
    sharedFolderChoices,
    sharedFoldersLoadError,
    noSharedFoldersForUser,
    selectedTargetFolderUid: '_default_',
    recordTitle: prefillTitle?.trim() || '',
    recordNotes: prefillNotes?.trim() || '',
  });

  await context.send({
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: formCard,
    }],
  });
}

/**
 * Route a command to the appropriate handler
 * @param {Object} context - Teams turn context
 * @param {string} text - Raw message text
 * @returns {Promise<boolean>} - True if command was handled, false otherwise
 */
async function routeCommand(context, text) {
  const { command, argsText } = parseCommand(text);
  
  switch (command) {
    case 'keeper-request-record':
    case 'request-record':
    case 'requestrecord':
    case 'rr':
      await handleRequestRecord(context, argsText);
      return true;
      
    case 'keeper-request-folder':
    case 'request-folder':
    case 'requestfolder':
    case 'rf':
      await handleRequestFolder(context, argsText);
      return true;
      
    case 'keeper-one-time-share':
    case 'share':
    case 'one-time-share':
    case 'onetimeshare':
    case 'ots':
      await handleShare(context, argsText);
      return true;
      
    case 'search-records':
    case 'searchrecords':
    case 'sr':
      await handleSearch(context, argsText, 'record');
      return true;
      
    case 'search-folders':
    case 'searchfolders':
    case 'sf':
      await handleSearch(context, argsText, 'folder');
      return true;
      
    case 'help':
    case 'h':
    case '?':
      await handleHelp(context);
      return true;
      
    case 'status':
    case 'st':
      await handleStatus(context);
      return true;

    case 'keeper-create-secret':
    case 'create-secret':
    case 'createsecret':
    case 'kcs':
      await handleCreateSecret(context, argsText);
      return true;
      
    default:
      return false;
  }
}

module.exports = {
  routeCommand,
  handleRequestRecord,
  handleRequestFolder,
  handleShare,
  handleSearch,
  handleHelp,
  handleStatus,
  handleCreateSecret,
  parseCommand,
  parseUidAndJustification,
  generateApprovalId,
  getUserInfo,
};
