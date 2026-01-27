/**
 * Configuration for Keeper Teams App
 * 
 * Configuration is loaded from environment variables.
 * Required variables should be set in .env file or environment.
 */

const config = {
  // ==================== Teams/Bot Configuration ====================
  MicrosoftAppId: process.env.CLIENT_ID,
  MicrosoftAppType: process.env.BOT_TYPE,
  MicrosoftAppTenantId: process.env.TENANT_ID,
  MicrosoftAppPassword: process.env.CLIENT_PASSWORD,

  // ==================== Keeper Configuration ====================
  keeper: {
    /**
     * URL of Keeper Commander Service Mode server
     * Example: http://localhost:3001/api/v2/
     */
    serviceUrl: process.env.KEEPER_SERVICE_URL || 'http://localhost:3001/api/v2/',
    
    /**
     * API key for authenticating with Service Mode (if required)
     */
    apiKey: process.env.KEEPER_API_KEY || null,
  },

  // ==================== Teams Channel Configuration ====================
  teams: {
    /**
     * Channel ID where approval requests are posted
     * This should be a dedicated approvals channel
     * 
     * HOW TO GET:
     * 1. Right-click the channel in Teams → "Get link to channel"
     * 2. The channel ID is in the URL after "channel/"
     * 3. Or send "/channel-status" to the bot in the channel to capture it
     * 
     * Example: 19:abc123def456@thread.tacv2
     */
    approvalsChannelId: process.env.APPROVALS_CHANNEL_ID || null,
    
    /**
     * Team ID containing the approvals channel (needed for proactive messages)
     * 
     * HOW TO GET:
     * 1. Right-click the team in Teams → "Get link to team"
     * 2. The team ID is the "groupId" parameter in the URL
     * 
     * Example: 12345678-1234-1234-1234-123456789abc
     */
    approvalsTeamId: process.env.APPROVALS_TEAM_ID || null,
  },

  // ==================== PEDM Configuration ====================
  pedm: {
    /**
     * Whether PEDM polling is enabled
     */
    enabled: process.env.PEDM_ENABLED === 'true',
    
    /**
     * Polling interval in milliseconds (default: 2 minutes)
     */
    pollingInterval: parseInt(process.env.PEDM_POLLING_INTERVAL) || 120000,
  },

  // ==================== Device Approval Configuration ====================
  deviceApproval: {
    /**
     * Whether device approval polling is enabled
     */
    enabled: process.env.DEVICE_APPROVAL_ENABLED === 'true',
    
    /**
     * Polling interval in milliseconds (default: 2 minutes)
     */
    pollingInterval: parseInt(process.env.DEVICE_APPROVAL_POLLING_INTERVAL) || 120000,
  },

  // ==================== Feature Flags ====================
  features: {
    /**
     * Require approval for one-time shares (vs direct creation)
     */
    requireShareApproval: process.env.REQUIRE_SHARE_APPROVAL === 'true',
    
    /**
     * Enable search commands
     */
    enableSearch: process.env.ENABLE_SEARCH !== 'false', // Enabled by default
  },
};

module.exports = config;
