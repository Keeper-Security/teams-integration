/**
 * Channel Service
 * 
 * Handles proactive messaging to Teams channels.
 * Manages conversation references for sending approval cards to dedicated channels.
 * 
 * Similar to Slack's approach of posting to approvals_channel_id,
 * this service routes approval requests to a dedicated Teams channel.
 */

const fs = require('fs');
const path = require('path');
const { MicrosoftAppCredentials, ConnectorClient } = require('botframework-connector');
const config = require('../config');

// ==================== Persistent Storage ====================

const DATA_DIR = path.join(__dirname, '..', 'data');
const REFERENCES_FILE = path.join(DATA_DIR, 'conversationReferences.json');

/**
 * Ensure data directory exists
 */
function ensureDataDir() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
    console.log('[ChannelService] Created data directory');
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
      console.log(`[ChannelService] Loaded ${Object.keys(expanded).length} conversation references from file`);
      return expanded;
    }
  } catch (error) {
    console.error('[ChannelService] Error loading references from file:', error.message);
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
    console.log('[ChannelService] Saved conversation references to file');
  } catch (error) {
    console.error('[ChannelService] Error saving references to file:', error.message);
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
 * In-memory store for approval card activity IDs
 * Maps approvalId -> activityId for updating cards later
 */
const approvalActivityMap = new Map();

/**
 * Store a conversation reference (with file persistence)
 * @param {string} key - Unique key (e.g., 'approvals', channel ID)
 * @param {Object} reference - Bot Framework conversation reference
 */
function storeConversationReference(key, reference) {
  conversationReferences.set(key, reference);
  console.log(`[ChannelService] Stored conversation reference for: ${key}`);
  
  // Persist to file (only for important keys like 'approvals')
  if (key === 'approvals' || key.startsWith('user:')) {
    const allRefs = Object.fromEntries(conversationReferences);
    saveReferencesToFile(allRefs);
  }
}

/**
 * Store an approval card's activity ID for later updates
 * @param {string} approvalId - The approval request ID
 * @param {string} activityId - The Teams activity ID of the card message
 */
function storeApprovalActivityId(approvalId, activityId) {
  approvalActivityMap.set(approvalId, activityId);
  console.log(`[ChannelService] Stored activity ID for approval ${approvalId}: ${activityId}`);
}

/**
 * Get the activity ID for an approval card
 * @param {string} approvalId - The approval request ID
 * @returns {string|null} - The activity ID or null
 */
function getApprovalActivityId(approvalId) {
  return approvalActivityMap.get(approvalId) || null;
}

/**
 * Remove an approval activity mapping (after it's been processed)
 * @param {string} approvalId - The approval request ID
 */
function removeApprovalActivityId(approvalId) {
  approvalActivityMap.delete(approvalId);
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
        console.log('[ChannelService] Captured APPROVALS channel reference');
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
      console.log(`[ChannelService] Captured user reference: ${userId}`);
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
      console.warn('[ChannelService] No approvals channel reference stored.');
      console.warn('[ChannelService] Send a message to the approvals channel to initialize.');
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
      console.error('[ChannelService] App not initialized');
      return false;
    }

    if (!reference || !reference.conversation || !reference.conversation.id) {
      console.error('[ChannelService] Invalid conversation reference');
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
      
      console.log('[ChannelService] Message sent successfully via app.send()');
      return true;
    } catch (error) {
      console.error('[ChannelService] Error sending message:', error.message);
      console.error('[ChannelService] Error details:', error.stack);
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
   * Then use Bot Connector Client to get the activity ID for later updates
   * @param {Object} card - Adaptive Card object
   * @param {string} approvalId - The approval request ID (for storing activity ID mapping)
   * @param {string} [preText] - Optional text to show before card
   * @returns {Promise<{success: boolean, activityId: string|null}>}
   */
  async sendApprovalCardViaConnector(card, approvalId, preText = null) {
    const approvalsRef = getConversationReference('approvals');
    
    if (!approvalsRef) {
      console.warn('[ChannelService] No approvals channel reference stored.');
      return { success: false, activityId: null };
    }

    if (!this.app) {
      console.error('[ChannelService] App not initialized');
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
      const conversationId = approvalsRef.conversation.id;
      const response = await this.app.send(conversationId, activity);
      
      // Log the response to see what we get back
      console.log('[ChannelService] app.send() response:', JSON.stringify(response));
      console.log('[ChannelService] app.send() response type:', typeof response);
      
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
        
        // Fallback: Extract messageid from conversation.id
        // Format: "19:xxx@thread.tacv2;messageid=1770017086743"
        if (!activityId && response.conversation?.id) {
          const match = response.conversation.id.match(/messageid=(\d+)/);
          if (match) {
            activityId = match[1];
            console.log('[ChannelService] Extracted activityId from conversation.id:', activityId);
          }
        }
      }
      
      console.log('[ChannelService] Final extracted activityId:', activityId);
      
      // Store the mapping of approvalId -> activityId for later updates
      if (approvalId && activityId) {
        storeApprovalActivityId(approvalId, activityId);
        console.log(`[ChannelService] Stored activity mapping: ${approvalId} -> ${activityId}`);
      } else {
        console.warn('[ChannelService] Could not store activity mapping - no activityId returned');
      }
      
      return { success: true, activityId: activityId };
    } catch (error) {
      console.error('[ChannelService] Error sending via app.send():', error.message);
      return { success: false, activityId: null };
    }
  }

  /**
   * Update an activity using Teams SDK API
   * @param {string} activityId - The activity ID to update
   * @param {Object} updatedCard - The updated Adaptive Card
   * @returns {Promise<boolean>}
   */
  async updateApprovalCard(activityId, updatedCard) {
    const approvalsRef = getConversationReference('approvals');
    
    if (!approvalsRef || !approvalsRef.serviceUrl) {
      console.error('[ChannelService] Missing approvals reference or serviceUrl');
      return false;
    }

    if (!this.app) {
      console.error('[ChannelService] App not initialized for update');
      return false;
    }

    // Get the base conversation ID (without messageid suffix)
    let conversationId = approvalsRef.conversation.id;
    // Remove any ";messageid=xxx" suffix if present
    if (conversationId.includes(';messageid=')) {
      conversationId = conversationId.split(';messageid=')[0];
    }

    console.log('[ChannelService] Attempting update:', {
      serviceUrl: approvalsRef.serviceUrl,
      conversationId,
      activityId,
    });

    try {
      // Build the updated activity
      const activity = {
        type: 'message',
        id: activityId,
        attachments: [{
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: updatedCard,
        }],
      };

      // Try using the app's API client first
      // The API client should use the same auth mechanism as app.send()
      if (this.app.api && this.app.api.conversations) {
        console.log('[ChannelService] Using app.api.conversations for update');
        const result = await this.app.api.conversations.activities(conversationId).update(
          activityId,
          activity
        );
        console.log('[ChannelService] Activity updated via Teams SDK API, activityId:', activityId);
        console.log('[ChannelService] Update result:', JSON.stringify(result));
        return true;
      }
      
      // Fallback: Try using the app's client directly
      // Create an API client with the correct service URL and the app's HTTP client
      console.log('[ChannelService] app.api.conversations not available, trying direct client');
      const { Client: ApiClient } = require('@microsoft/teams.api');
      const apiClient = new ApiClient(approvalsRef.serviceUrl, this.app.client);
      
      const result = await apiClient.conversations.activities(conversationId).update(
        activityId,
        activity
      );
      
      console.log('[ChannelService] Activity updated via direct API client, activityId:', activityId);
      console.log('[ChannelService] Update result:', JSON.stringify(result));
      return true;
    } catch (error) {
      console.error('[ChannelService] Error updating activity:', error.message);
      console.error('[ChannelService] Error stack:', error.stack);
      
      // Last resort: Try using the ConnectorClient
      try {
        console.log('[ChannelService] Trying ConnectorClient as last resort');
        const client = getConnectorClient(approvalsRef.serviceUrl);
        
        const fullActivity = {
          type: 'message',
          id: activityId,
          from: approvalsRef.bot,
          conversation: { id: conversationId },
          channelId: approvalsRef.channelId,
          serviceUrl: approvalsRef.serviceUrl,
          attachments: [{
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: updatedCard,
          }],
        };

        await client.conversations.updateActivity(
          conversationId,
          activityId,
          fullActivity
        );
        
        console.log('[ChannelService] Activity updated via ConnectorClient (last resort)');
        return true;
      } catch (connectorError) {
        console.error('[ChannelService] ConnectorClient also failed:', connectorError.message);
        return false;
      }
    }
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
      console.warn(`[ChannelService] No reference for user: ${userId}`);
      return false;
    }

    return this.sendToChannel(userRef, message);
  }

  /**
   * Check if approvals channel is configured and initialized
   * @returns {boolean}
   */
  isApprovalsChannelReady() {
    const hasConfig = !!config.teams?.approvalsChannelId;
    const hasReference = !!getConversationReference('approvals');
    return hasConfig && hasReference;
  }

  /**
   * Get status information
   * @returns {Object}
   */
  getStatus() {
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
    console.log('[ChannelService] Initialized');
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
};
