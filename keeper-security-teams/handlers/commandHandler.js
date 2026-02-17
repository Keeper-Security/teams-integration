/**
 * Command Handler for Keeper Security Bot
 * 
 * Handles commands from users:
 * - keeper-request-record <name> <justification>
 * - keeper-request-folder <name> <justification>
 * - keeper-one-time-share <name> [justification]
 * - help
 */

const keeperClient = require('../services/keeperClient');
const { getChannelService } = require('../services');
const graphService = require('../services/graphService');
const cards = require('../cards');
const config = require('../config');
const { isUid } = require('../utils/helpers');

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
  if (trimmed.startsWith('"') || trimmed.startsWith("'")) {
    const quote = trimmed[0];
    const endQuote = trimmed.indexOf(quote, 1);
    if (endQuote > 0) {
      const uid = trimmed.substring(1, endQuote);
      const justification = trimmed.substring(endQuote + 1).trim();
      return { uid, justification };
    }
  }
  
  // Otherwise, first word is uid, rest is justification
  const [uid, ...rest] = trimmed.split(/\s+/);
  return { uid, justification: rest.join(' ') };
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
        console.log(`[getUserInfo] Fetched email for user ${from.name}: ${email}`);
      } else {
        console.warn(`[getUserInfo] No email found for user ${from.name} (AAD ID: ${aadObjectId})`);
      }
    } catch (error) {
      console.error(`[getUserInfo] Error fetching email for user ${from.name}:`, error.message);
      // Fallback: check if email is directly in activity (rare, but possible)
      userInfo.userEmail = from.email || from.userPrincipalName || null;
    }
  } else {
    // No AAD Object ID available - try fallback fields
    userInfo.userEmail = from.email || from.userPrincipalName || null;
    if (!userInfo.userEmail) {
      console.warn(`[getUserInfo] No AAD Object ID found for user ${from.name}, cannot fetch email`);
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
  
  if (!uid) {
    await context.send('❌ **Usage:** `keeper-request-record <record-name> <justification>`\n\nExample: `keeper-request-record AWS-Prod Need for deployment`');
    return;
  }
  
  if (!justification) {
    await context.send('❌ Please provide a justification for your access request.\n\n**Usage:** `keeper-request-record <record-name> <justification>`');
    return;
  }
  
  // Check if identifier is a UID or description
  const isUidFormat = isUid(uid.trim());
  let record = null;
  let recordUid = null;
  let recordTitle = uid; // Default to the identifier
  
  if (isUidFormat) {
    // Try to get record by UID
    record = await keeperClient.getRecordByUid(uid.trim());
    if (record) {
      recordUid = record.uid;
      recordTitle = record.title;
    }
  }
  
  // If not found by UID or not a UID format, treat as description
  if (!record && isUidFormat) {
    // UID format but not found - might be invalid
    await context.send('❌ Record not found: `' + uid + '`\n\nPlease check the UID and try again.');
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
        `📋 New record access request from **${userInfo.userName}**`
      );
      sentToChannel = result.success;
      
      if (result.activityId) {
        console.log(`[CommandHandler] Approval card sent with activityId: ${result.activityId}`);
      }
    } else {
      console.log('[CommandHandler] Cannot send to approvals channel:', {
        approvalsReady: status.approvalsChannelReady,
        appReady: status.appReady
      });
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
  
  if (!uid) {
    await context.send('❌ **Usage:** `keeper-request-folder <folder-name> <justification>`\n\nExample: `keeper-request-folder "Engineering Creds" Project onboarding`');
    return;
  }
  
  if (!justification) {
    await context.send('❌ Please provide a justification for your access request.\n\n**Usage:** `keeper-request-folder <folder-name> <justification>`');
    return;
  }
  
  // Check if identifier is a UID or description
  const isUidFormat = isUid(uid.trim());
  let folder = null;
  let folderUid = null;
  let folderName = uid; // Default to the identifier
  
  if (isUidFormat) {
    // Try to get folder by UID
    folder = await keeperClient.getFolderByUid(uid.trim());
    if (folder) {
      folderUid = folder.uid;
      folderName = folder.name;
    }
  }
  
  // If not found by UID or not a UID format, treat as description
  if (!folder && isUidFormat) {
    // UID format but not found - might be invalid
    await context.send('❌ Folder not found: `' + uid + '`\n\nPlease check the UID and try again.');
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
        `📁 New folder access request from **${userInfo.userName}**`
      );
      sentToChannel = result.success;
      
      if (result.activityId) {
        console.log(`[CommandHandler] Folder approval card sent with activityId: ${result.activityId}`);
      }
    } else {
      console.log('[CommandHandler] Cannot send to approvals channel:', {
        approvalsReady: status.approvalsChannelReady,
        appReady: status.appReady
      });
    }
  }
  
  if (sentToChannel) {
    // Approval sent to dedicated channel - notify requester
    // Show the identifier (UID) the user entered, not the resolved name (for security)
    // Using two spaces before \n for Markdown line breaks
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
  
  if (!uid) {
    await context.send('**Usage:** `keeper-one-time-share <record-name> [justification]`\n\nExample: `keeper-one-time-share AWS-Prod Share with contractor`');
    return;
  }
  
  if (!justification) {
    await context.send('**Justification is required.**\n\n**Usage:** `keeper-one-time-share "<record-name>" <justification>`\n\nExample: `keeper-one-time-share "AWS-Prod" Share with contractor for audit`');
    return;
  }
  
  // Check if input is a valid UID (22 char base64)
  const inputIsUid = isUid(uid);
  const requiresApproval = config.features.requireShareApproval;
  
  // For approval flow with description, we don't need to fetch record details upfront
  // The admin will search and select the correct record
  let record = null;
  let recordUid = null;
  let recordTitle = uid;
  
  // Only fetch record details if it's a UID (for validation and display)
  if (inputIsUid) {
    record = await keeperClient.getRecordByUid(uid);
    if (record) {
      recordUid = record.uid;
      recordTitle = record.title;
      
      // Check if it's a folder (one-time-share only works for records)
      if (record.record_type && (record.record_type.includes('folder') || record.record_type === 'shared_folder')) {
        await context.send('**One-time share is only available for records, not folders.**\n\nPlease use `keeper-request-folder` for folder access.');
        return;
      }
    } else {
      await context.send('Record not found with UID: `' + uid + '`\n\nPlease verify the UID and try again.');
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
          console.log(`[CommandHandler] Share approval card sent with activityId: ${sendResult.activityId}`);
        }
      } else {
        console.log('[CommandHandler] Cannot send to approvals channel:', {
          approvalsReady: status.approvalsChannelReady,
          appReady: status.appReady
        });
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
    await context.send('❌ **Usage:** `search-' + typeLabel + ' <query>`\n\nExample: `search-' + typeLabel + ' production`');
    return;
  }
  
  if (!config.features.enableSearch) {
    await context.send('❌ Search is not enabled. Contact your administrator.');
    return;
  }
  
  await context.send('🔍 Searching...');
  
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
  await context.send('🔄 Checking Keeper connection...');
  
  const isHealthy = await keeperClient.healthCheck();
  
  if (isHealthy) {
    await context.send('✅ **Keeper Status:** Connected\n\n' +
      '• Service URL: `' + (config.keeper?.serviceUrl || 'Not configured') + '`\n' +
      '• EPM Polling: ' + (config.pedm?.enabled ? '✅ Enabled' : '❌ Disabled') + '\n' +
      '• Device Approval: ' + (config.deviceApproval?.enabled ? '✅ Enabled' : '❌ Disabled') + '\n' +
      '• Search: ' + (config.features?.enableSearch ? '✅ Enabled' : '❌ Disabled')
    );
  } else {
    await context.send('❌ **Keeper Status:** Disconnected\n\n' +
      'Unable to connect to Keeper Commander Service Mode.\n' +
      'Service URL: `' + (config.keeper?.serviceUrl || 'Not configured') + '`\n\n' +
      'Please check that Keeper Commander is running in Service Mode.'
    );
  }
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
  parseCommand,
  parseUidAndJustification,
  generateApprovalId,
  getUserInfo,
};
