/**
 * Channel Service
 * 
 * Handles proactive messaging to Teams channels.
 * Manages conversation references for sending approval cards to dedicated channels.
 * 
 * Posts approval cards to the configured approvals channel.
 * this service routes approval requests to a dedicated Teams channel.
 */

const fs = require('fs');
const path = require('path');
const { MicrosoftAppCredentials, ConnectorClient } = require('botframework-connector');
const { getConfig } = require('../config');
const { createLogger } = require('./logger');

const log = createLogger('ChannelService');

// ==================== Persistent Storage ====================

const DATA_DIR = path.join(__dirname, '..', 'data');
const REFERENCES_FILE = path.join(DATA_DIR, 'conversationReferences.json');
const APPROVAL_STATUS_FILE = path.join(DATA_DIR, 'approvalStatus.json');
const ACTIVITY_IDS_FILE = path.join(DATA_DIR, 'activityIds.json');

// Max age for approval status entries (7 days) - auto-cleanup old entries
const APPROVAL_STATUS_MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000;

/**
 * Ensure data directory exists
 */
function ensureDataDir() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
    log.info('Created data directory');
  }
}

/**
 * Load conversation references from file (and expand if minimized)
 */
function loadReferencesFromFile() {
  try {
    ensureDataDir();
    if (fs.existsSync(REFERENCES_FILE)) {
      const data = fs.readFileSync(REFERENCES_FILE, 'utf8');
      const refs = JSON.parse(data);
      // Expand minimized references
      const expanded = {};
      for (const [key, ref] of Object.entries(refs)) {
        expanded[key] = expandReference(ref);
      }
      log.info(`Loaded ${Object.keys(expanded).length} conversation references from file`);
      return expanded;
    }
  } catch (error) {
    log.error('Error loading references from file', error.message);
  }
  return {};
}

/**
 * Minimize a conversation reference to essential fields only
 * @param {Object} ref - Full conversation reference
 * @returns {Object} - Minimized reference
 */
function minimizeReference(ref) {
  if (!ref) return ref;
  return {
    serviceUrl: ref.serviceUrl,
    conversationId: ref.conversation?.id,
    conversationType: ref.conversation?.conversationType,
    tenantId: ref.conversation?.tenantId,
    botId: ref.bot?.id
  };
}

/**
 * Expand a minimized reference back to full format
 * @param {Object} min - Minimized reference
 * @returns {Object} - Full conversation reference
 */
function expandReference(min) {
  if (!min) return min;
  // If already in full format, return as-is
  if (min.conversation) return min;
  return {
    serviceUrl: min.serviceUrl,
    channelId: 'msteams',
    conversation: {
      id: min.conversationId,
      conversationType: min.conversationType,
      tenantId: min.tenantId
    },
    bot: {
      id: min.botId
    }
  };
}

/**
 * Save conversation references to file (minimized)
 */
function saveReferencesToFile(references) {
  try {
    ensureDataDir();
    // Minimize each reference before saving
    const minimized = {};
    for (const [key, ref] of Object.entries(references)) {
      minimized[key] = minimizeReference(ref);
    }
    const data = JSON.stringify(minimized, null, 2);
    fs.writeFileSync(REFERENCES_FILE, data, 'utf8');
    log.debug('Saved conversation references to file');
  } catch (error) {
    log.error('Error saving references to file', error.message);
  }
}

/**
 * Load approval status from file
 * Also cleans up entries older than APPROVAL_STATUS_MAX_AGE_MS
 */
function loadApprovalStatusFromFile() {
  try {
    ensureDataDir();
    if (fs.existsSync(APPROVAL_STATUS_FILE)) {
      const data = fs.readFileSync(APPROVAL_STATUS_FILE, 'utf8');
      const statuses = JSON.parse(data);
      const now = Date.now();
      const cleaned = {};
      let cleanedCount = 0;
      
      // Filter out old entries
      for (const [key, status] of Object.entries(statuses)) {
        const updatedAt = new Date(status.updatedAt).getTime();
        if (now - updatedAt < APPROVAL_STATUS_MAX_AGE_MS) {
          cleaned[key] = status;
        } else {
          cleanedCount++;
        }
      }
      
      if (cleanedCount > 0) {
        log.info(`Cleaned up ${cleanedCount} old approval status entries`);
        // Save cleaned data back to file
        saveApprovalStatusToFile(cleaned);
      }
      
      log.info(`Loaded ${Object.keys(cleaned).length} approval status entries from file`);
      return cleaned;
    }
  } catch (error) {
    log.error('Error loading approval status from file', error.message);
  }
  return {};
}

/**
 * Save approval status to file
 */
function saveApprovalStatusToFile(statuses) {
  try {
    ensureDataDir();
    const data = JSON.stringify(statuses, null, 2);
    fs.writeFileSync(APPROVAL_STATUS_FILE, data, 'utf8');
    log.debug('Saved approval status to file');
  } catch (error) {
    log.error('Error saving approval status to file', error.message);
  }
}

/**
 * Load activity IDs from file
 */
function loadActivityIdsFromFile() {
  try {
    ensureDataDir();
    if (fs.existsSync(ACTIVITY_IDS_FILE)) {
      const data = fs.readFileSync(ACTIVITY_IDS_FILE, 'utf8');
      const ids = JSON.parse(data);
      log.info(`Loaded ${Object.keys(ids).length} activity IDs from file`);
      return ids;
    }
  } catch (error) {
    log.error('Error loading activity IDs from file', error.message);
  }
  return {};
}

/**
 * Save activity IDs to file
 */
function saveActivityIdsToFile(ids) {
  try {
    ensureDataDir();
    const data = JSON.stringify(ids, null, 2);
    fs.writeFileSync(ACTIVITY_IDS_FILE, data, 'utf8');
    log.debug('Saved activity IDs to file');
  } catch (error) {
    log.error('Error saving activity IDs to file', error.message);
  }
}

// Cache for connector clients by service URL
const connectorClients = new Map();

/**
 * Get or create a ConnectorClient for the given service URL
 * @param {string} serviceUrl - The Bot Framework service URL
 * @returns {ConnectorClient}
 */
function getConnectorClient(serviceUrl) {
  if (!connectorClients.has(serviceUrl)) {
    const config = getConfig();
    
    // Trust the service URL (required for Teams)
    MicrosoftAppCredentials.trustServiceUrl(serviceUrl);
    
    // Create credentials
    const credentials = new MicrosoftAppCredentials(
      config.MicrosoftAppId,
      config.MicrosoftAppPassword
    );
    
    // Create connector client
    const client = new ConnectorClient(credentials, { baseUri: serviceUrl });
    connectorClients.set(serviceUrl, client);
  }
  
  return connectorClients.get(serviceUrl);
}

/**
 * Conversation references store with file persistence
 * Loaded from file on startup, saved on updates
 */
const conversationReferences = new Map(Object.entries(loadReferencesFromFile()));

/**
 * Persistent store for approval card activity IDs (loaded from file on startup)
 * Maps approvalId -> activityId for updating cards later
 */
const approvalActivityMap = new Map(Object.entries(loadActivityIdsFromFile()));

/**
 * Persistent store for approval status (loaded from file on startup)
 * Maps approvalId -> { status, approverName, permission, duration, expiresAt, processedTime, ... }
 * This is used by the refresh mechanism to return the correct card state
 */
const approvalStatusMap = new Map(Object.entries(loadApprovalStatusFromFile()));

/**
 * Store approval status for refresh mechanism (persisted to file)
 * @param {string} approvalId - The approval request ID
 * @param {Object} statusData - Status data including status, approverName, etc.
 */
function storeApprovalStatus(approvalId, statusData) {
  approvalStatusMap.set(approvalId, {
    ...statusData,
    updatedAt: new Date().toISOString(),
  });
  log.debug(`Stored approval status for ${approvalId}: ${statusData.status}`);
  
  // Persist to file
  const allStatuses = Object.fromEntries(approvalStatusMap);
  saveApprovalStatusToFile(allStatuses);
}

/**
 * Get approval status
 * @param {string} approvalId - The approval request ID
 * @returns {Object|null} - Status data or null if not processed
 */
function getApprovalStatus(approvalId) {
  return approvalStatusMap.get(approvalId) || null;
}

/**
 * Check if an approval has been processed
 * @param {string} approvalId - The approval request ID
 * @returns {boolean}
 */
function isApprovalProcessed(approvalId) {
  return approvalStatusMap.has(approvalId);
}

/**
 * Store a conversation reference (with file persistence)
 * @param {string} key - Unique key (e.g., 'approvals', channel ID)
 * @param {Object} reference - Bot Framework conversation reference
 */
function storeConversationReference(key, reference) {
  conversationReferences.set(key, reference);
  log.debug(`Stored conversation reference for: ${key}`);
  
  // Persist to file (only for important keys like 'approvals')
  if (key === 'approvals' || key.startsWith('user:')) {
    const allRefs = Object.fromEntries(conversationReferences);
    saveReferencesToFile(allRefs);
  }
}

/**
 * Store an approval card's activity ID for later updates (persisted to file)
 * @param {string} approvalId - The approval request ID
 * @param {string} activityId - The Teams activity ID of the card message
 */
function storeApprovalActivityId(approvalId, activityId) {
  approvalActivityMap.set(approvalId, activityId);
  log.info(`Stored activity ID for approval ${approvalId}: ${activityId}`);
  
  // Persist to file
  const allIds = Object.fromEntries(approvalActivityMap);
  saveActivityIdsToFile(allIds);
}

/**
 * Get the activity ID for an approval card
 * @param {string} approvalId - The approval request ID
 * @returns {string|null} - The activity ID or null
 */
function getApprovalActivityId(approvalId) {
  const activityId = approvalActivityMap.get(approvalId) || null;
  log.debug(`Retrieved activity ID for approval ${approvalId}: ${activityId}`);
  return activityId;
}

/**
 * Remove an approval activity mapping (after it's been processed)
 * @param {string} approvalId - The approval request ID
 */
function removeApprovalActivityId(approvalId) {
  approvalActivityMap.delete(approvalId);
  log.debug(`Removed activity ID for approval ${approvalId}`);
  
  // Persist to file
  const allIds = Object.fromEntries(approvalActivityMap);
  saveActivityIdsToFile(allIds);
}

/**
 * Get a stored conversation reference
 * @param {string} key - Unique key
 * @returns {Object|null} - Conversation reference or null
 */
function getConversationReference(key) {
  return conversationReferences.get(key) || null;
}

/**
 * Extract and store conversation reference from a Teams activity
 * @param {Object} activity - Teams activity object
 * @returns {Object} - Conversation reference
 */
function extractConversationReference(activity) {
  return {
    activityId: activity.id,
    user: activity.from,
    bot: activity.recipient,
    conversation: activity.conversation,
    channelId: activity.channelId,
    locale: activity.locale,
    serviceUrl: activity.serviceUrl,
  };
}

/**
 * Check if an activity is from the configured approvals channel
 * @param {Object} activity - Teams activity object
 * @returns {boolean}
 */
function isApprovalsChannel(activity) {
  const config = getConfig();
  const approvalsChannelId = config.teams?.approvalsChannelId;
  
  if (!approvalsChannelId) {
    return false;
  }
  
  // Teams channel ID can be in conversation.id or channelData.teamsChannelId
  const channelId = activity.channelData?.teamsChannelId || activity.conversation?.id;
  
  // Check if it matches (partial match for Teams channel format)
  return channelId && (
    channelId === approvalsChannelId ||
    channelId.includes(approvalsChannelId)
  );
}

/**
 * Check if activity is from a Teams channel (not 1:1 chat)
 * @param {Object} activity - Teams activity object
 * @returns {boolean}
 */
function isTeamsChannel(activity) {
  const conversationType = activity.conversation?.conversationType;
  return conversationType === 'channel' || 
         activity.channelData?.teamsChannelId != null;
}

/**
 * ChannelService class - manages proactive messaging to Teams channels
 */
class ChannelService {
  constructor(app) {
    this.app = app;
    const config = getConfig();
    this.appId = process.env.CLIENT_ID || config.MicrosoftAppId;
  }

  /**
   * Store conversation reference from current context
   * @param {Object} context - Teams turn context
   * @param {string} key - Storage key
   */
  captureConversationReference(context, key = null) {
    const activity = context.activity;
    const reference = extractConversationReference(activity);
    
    // If no key provided, try to determine channel type
    if (!key) {
      if (isApprovalsChannel(activity)) {
        key = 'approvals';
        log.info('Captured approvals channel reference');
      } else if (isTeamsChannel(activity)) {
        key = activity.channelData?.teamsChannelId || activity.conversation?.id;
      } else {
        key = activity.conversation?.id || 'default';
      }
    }
    
    storeConversationReference(key, reference);
    return reference;
  }

  /**
   * Store a 1:1 conversation reference for a user (for DMs)
   * @param {Object} context - Teams turn context
   */
  captureUserReference(context) {
    const activity = context.activity;
    const userId = activity.from?.id;
    
    if (userId && !isTeamsChannel(activity)) {
      const reference = extractConversationReference(activity);
      storeConversationReference(`user:${userId}`, reference);
      log.debug(`Captured user reference: ${userId}`);
    }
  }

  /**
   * Send a message to the approvals channel
   * @param {Object} message - Message to send (card or text)
   * @returns {Promise<boolean>} - Success status
   */
  async sendToApprovalsChannel(message) {
    const approvalsRef = getConversationReference('approvals');
    
    if (!approvalsRef) {
      log.warn('No approvals channel reference stored. Send a message to the approvals channel to initialize.');
      return false;
    }

    return this.sendToChannel(approvalsRef, message);
  }

  /**
   * Send message to a specific channel using conversation reference
   * Uses the App.send() method from @microsoft/teams.apps SDK
   * @param {Object} reference - Conversation reference
   * @param {Object} message - Message to send
   * @returns {Promise<boolean>} - Success status
   */
  async sendToChannel(reference, message) {
    if (!this.app) {
      log.error('App not initialized');
      return false;
    }

    if (!reference || !reference.conversation || !reference.conversation.id) {
      log.error('Invalid conversation reference');
      return false;
    }

    try {
      // Prepare the activity to send
      let activity;
      
      if (typeof message === 'string') {
        activity = { type: 'message', text: message };
      } else if (message.type === 'message') {
        // Already formatted as an activity
        activity = message;
      } else if (message.attachments) {
        // Has attachments, use as-is
        activity = { type: 'message', ...message };
      } else {
        // Assume it's a card object, wrap it
        activity = {
          type: 'message',
          attachments: [{
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: message,
          }],
        };
      }

      // Use the App.send() method - this is the correct API for @microsoft/teams.apps v2.0+
      // conversationId is the conversation.id from the stored reference
      const conversationId = reference.conversation.id;
      
      await this.app.send(conversationId, activity);
      
      log.debug('Message sent successfully via app.send()');
      return true;
    } catch (error) {
      log.error('Error sending message', { message: error.message, stack: error.stack });
      return false;
    }
  }

  /**
   * Send an Adaptive Card to the approvals channel
   * @param {Object} card - Adaptive Card object
   * @param {string} [preText] - Optional text to show before card
   * @returns {Promise<boolean>}
   */
  async sendApprovalCard(card, preText = null) {
    const message = {
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: card,
      }],
    };

    if (preText) {
      message.text = preText;
    }

    return this.sendToApprovalsChannel(message);
  }

  /**
   * Send an Adaptive Card to the approvals channel using Teams SDK app.send()
   * Stores the activity ID mapping for later card updates
   * @param {Object} card - Adaptive Card object
   * @param {string} approvalId - The approval request ID (for storing activity ID mapping)
   * @param {string} [preText] - Optional text to show before card
   * @returns {Promise<{success: boolean, activityId: string|null}>}
   */
  async sendApprovalCardViaConnector(card, approvalId, preText = null) {
    const approvalsRef = getConversationReference('approvals');
    
    if (!approvalsRef) {
      log.warn('No approvals channel reference stored.');
      return { success: false, activityId: null };
    }

    if (!this.app) {
      log.error('App not initialized');
      return { success: false, activityId: null };
    }

    try {
      // Build the activity
      const activity = {
        type: 'message',
        attachments: [{
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: card,
        }],
      };

      if (preText) {
        activity.text = preText;
      }

      // Use Teams SDK app.send() for sending (this works for proactive messaging)
      const conversationId = approvalsRef.conversationId || approvalsRef.conversation?.id;
      const response = await this.app.send(conversationId, activity);
      
      log.debug('app.send() response', { response, type: typeof response, approvalId });
      
      // Try to extract activity ID from response
      let activityId = null;
      if (response) {
        if (typeof response === 'string') {
          activityId = response;
        } else if (response.id) {
          activityId = response.id;
        } else if (response.activityId) {
          activityId = response.activityId;
        }
        
        if (!activityId && response.conversation?.id) {
          const match = response.conversation.id.match(/messageid=(\d+)/);
          if (match) {
            activityId = match[1];
            log.debug('Extracted activityId from conversation.id', activityId);
          }
        }
      }
      
      log.debug('Final extracted activityId', { activityId, approvalId });
      
      if (approvalId && activityId) {
        storeApprovalActivityId(approvalId, activityId);
        log.info(`Stored activity mapping: ${approvalId} -> ${activityId}`);
      } else {
        log.warn('Could not store activity mapping - no activityId returned', { approvalId });
      }
      
      return { success: true, activityId: activityId };
    } catch (error) {
      log.error('Error sending via app.send()', { message: error.message, approvalId });
      return { success: false, activityId: null };
    }
  }

  /**
   * Update an activity using Bot Framework ConnectorClient
   * Uses the bot's own credentials for proper authorization
   * @param {string} activityId - The activity ID to update
   * @param {Object} updatedCard - The updated Adaptive Card
   * @returns {Promise<boolean>}
   */
  async updateApprovalCard(activityId, updatedCard) {
    const approvalsRef = getConversationReference('approvals');
    
    if (!approvalsRef || !approvalsRef.serviceUrl) {
      log.error('Missing approvals reference or serviceUrl');
      return false;
    }

    let conversationId = approvalsRef.conversationId || approvalsRef.conversation?.id;
    if (conversationId && conversationId.includes(';messageid=')) {
      conversationId = conversationId.split(';messageid=')[0];
    }

    if (!conversationId) {
      log.error('Missing conversation ID');
      return false;
    }

    log.debug('Attempting card update', { conversationId, activityId, serviceUrl: approvalsRef.serviceUrl });

    // Build the updated activity with all required fields
    const fullActivity = {
      type: 'message',
      id: activityId,
      from: approvalsRef.bot ? { id: approvalsRef.bot.id || approvalsRef.botId } : { id: approvalsRef.botId },
      conversation: { id: conversationId },
      channelId: approvalsRef.channelId || 'msteams',
      serviceUrl: approvalsRef.serviceUrl,
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: updatedCard,
      }],
    };

    // Use ConnectorClient with bot credentials (most reliable method)
    try {
      log.debug('Using ConnectorClient with bot credentials');
      const client = getConnectorClient(approvalsRef.serviceUrl);
      
      await client.conversations.updateActivity(
        conversationId,
        activityId,
        fullActivity
      );
      
      log.info('Activity updated successfully via ConnectorClient', { activityId });
      return true;
    } catch (connectorError) {
      log.error('ConnectorClient update failed', { 
        message: connectorError.message,
        statusCode: connectorError.response?.status,
        body: connectorError.response?.data
      });
    }

    // Fallback: Try Teams SDK API if available
    if (this.app?.api?.conversations) {
      try {
        log.debug('Fallback: Using app.api.conversations');
        const activity = {
          type: 'message',
          id: activityId,
          attachments: [{
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: updatedCard,
          }],
        };
        
        await this.app.api.conversations.activities(conversationId).update(activityId, activity);
        log.info('Activity updated via Teams SDK API fallback', { activityId });
        return true;
      } catch (apiError) {
        log.error('Teams SDK API update also failed', apiError.message);
      }
    }

    return false;
  }

  /**
   * Send a Direct Message (1:1) to a user
   * @param {string} userId - Teams user ID
   * @param {Object|string} message - Message to send
   * @returns {Promise<boolean>} - Success status
   */
  async sendDirectMessage(userId, message) {
    const userRef = getConversationReference(`user:${userId}`);
    
    if (!userRef) {
      log.warn(`No reference for user: ${userId}`);
      return false;
    }

    return this.sendToChannel(userRef, message);
  }

  /**
   * Check if approvals channel is configured and initialized
   * @returns {boolean}
   */
  isApprovalsChannelReady() {
    const config = getConfig();
    const hasConfig = !!config.teams?.approvalsChannelId;
    const hasReference = !!getConversationReference('approvals');
    return hasConfig && hasReference;
  }

  /**
   * Get status information
   * @returns {Object}
   */
  getStatus() {
    const config = getConfig();
    return {
      approvalsChannelConfigured: !!config.teams?.approvalsChannelId,
      approvalsChannelId: config.teams?.approvalsChannelId || 'Not configured',
      approvalsChannelReady: this.isApprovalsChannelReady(),
      appReady: !!this.app,
      storedReferences: Array.from(conversationReferences.keys()),
    };
  }
}

// Singleton instance
let channelServiceInstance = null;

/**
 * Initialize the channel service
 * @param {Object} app - Teams app instance
 * @returns {ChannelService}
 */
function initializeChannelService(app) {
  if (!channelServiceInstance) {
    channelServiceInstance = new ChannelService(app);
    log.info('Channel service initialized');
  }
  return channelServiceInstance;
}

/**
 * Get the channel service instance
 * @returns {ChannelService|null}
 */
function getChannelService() {
  return channelServiceInstance;
}

module.exports = {
  ChannelService,
  initializeChannelService,
  getChannelService,
  storeConversationReference,
  getConversationReference,
  extractConversationReference,
  isApprovalsChannel,
  isTeamsChannel,
  storeApprovalActivityId,
  getApprovalActivityId,
  removeApprovalActivityId,
  // Approval status storage for refresh mechanism
  storeApprovalStatus,
  getApprovalStatus,
  isApprovalProcessed,
};
