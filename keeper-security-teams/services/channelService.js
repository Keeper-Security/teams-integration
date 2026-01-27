/**
 * Channel Service
 * 
 * Handles proactive messaging to Teams channels.
 * Manages conversation references for sending approval cards to dedicated channels.
 * 
 * Similar to Slack's approach of posting to approvals_channel_id,
 * this service routes approval requests to a dedicated Teams channel.
 */

const config = require('../config');

/**
 * In-memory store for conversation references
 * In production, this should be persisted to a database
 */
const conversationReferences = new Map();

/**
 * Store a conversation reference
 * @param {string} key - Unique key (e.g., 'approvals', channel ID)
 * @param {Object} reference - Bot Framework conversation reference
 */
function storeConversationReference(key, reference) {
  conversationReferences.set(key, reference);
  console.log(`[ChannelService] Stored conversation reference for: ${key}`);
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
};
