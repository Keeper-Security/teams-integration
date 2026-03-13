/**
 * Configuration for Keeper Teams App
 */

const fs = require('fs');
const path = require('path');

let ksmService = null;
let configInitialized = false;
let cachedConfig = null;

function loadKsmService() {
  if (ksmService !== null) return ksmService;
  
  try {
    ksmService = require('./services/ksmService');
    return ksmService;
  } catch (error) {
    ksmService = false;
    return null;
  }
}

/**
 * Load configuration from KSM if available
 */
async function loadFromKsm() {
  const ksm = loadKsmService();
  if (!ksm) return null;
  
  // Check for KSM config sources
  let ksmConfig = null;
  let commanderRecord = null;
  let teamsRecord = null;
  
  // Check Docker secrets first
  const dockerKsmPath = '/run/secrets/ksm-config';
  const dockerCommanderPath = '/run/secrets/commander-record';
  const dockerTeamsPath = '/run/secrets/teams-record';
  
  if (fs.existsSync(dockerKsmPath)) {
    try {
      ksmConfig = fs.readFileSync(dockerKsmPath, 'utf8').trim();
    } catch (error) {
      console.warn('Failed to read KSM config from Docker secret:', error.message);
    }
  }
  
  if (fs.existsSync(dockerCommanderPath)) {
    try {
      commanderRecord = fs.readFileSync(dockerCommanderPath, 'utf8').trim();
    } catch (error) {
      console.warn('Failed to read commander record from Docker secret:', error.message);
    }
  }
  
  if (fs.existsSync(dockerTeamsPath)) {
    try {
      teamsRecord = fs.readFileSync(dockerTeamsPath, 'utf8').trim();
    } catch (error) {
      console.warn('Failed to read teams record from Docker secret:', error.message);
    }
  }
  
  // Fallback to environment variables
  if (!ksmConfig) ksmConfig = process.env.KSM_CONFIG;
  if (!commanderRecord) commanderRecord = process.env.COMMANDER_RECORD || 'CSMD config';
  if (!teamsRecord) teamsRecord = process.env.TEAMS_RECORD || 'CSMD teams config';
  
  // If no KSM config, skip KSM loading
  if (!ksmConfig) {
    return null;
  }
  
  try {
    const ksmData = await ksm.fetchCredentialsFromKsm(
      ksmConfig,
      commanderRecord,
      teamsRecord
    );
    return ksmData;
  } catch (error) {
    console.error('Failed to load from KSM:', error.message);
    return null;
  }
}

/**
 * Build configuration object from KSM data and environment variables
 * KSM values take precedence over environment variables
 */
function buildConfig(ksmData = null) {
  const ksm = ksmData || {};
  
  return {
    // ==================== Teams/Bot Configuration ====================
    MicrosoftAppId: ksm.teams?.clientId || process.env.CLIENT_ID,
    MicrosoftAppType: ksm.teams?.botType || process.env.BOT_TYPE || 'MultiTenant',
    MicrosoftAppTenantId: ksm.teams?.tenantId || process.env.TENANT_ID,
    MicrosoftAppPassword: ksm.teams?.clientSecret || process.env.CLIENT_SECRET,

    // ==================== Keeper Configuration ====================
    keeper: {
      serviceUrl: ksm.keeper?.serviceUrl || process.env.KEEPER_SERVICE_URL || 'http://localhost:8900/api/v2/',
      apiKey: ksm.keeper?.apiKey || process.env.KEEPER_API_KEY || null,
    },

    // ==================== Teams Channel Configuration ====================
    teams: {
      approvalsChannelId: ksm.teams?.approvalsChannelId || process.env.APPROVALS_CHANNEL_ID || null,
      approvalsTeamId: ksm.teams?.approvalsTeamId || process.env.APPROVALS_TEAM_ID || null,
    },

    // ==================== EPM (Endpoint Privilege Manager) Configuration ====================
    pedm: {
      enabled: ksm.pedm?.enabled ?? (process.env.EPM_ENABLED === 'true'),
      pollingInterval: ((ksm.pedm?.pollingInterval || parseInt(process.env.EPM_POLL_INTERVAL)) || 120) * 1000,
    },

    // ==================== Device Approval Configuration ====================
    deviceApproval: {
      enabled: ksm.deviceApproval?.enabled ?? (process.env.DEVICE_APPROVAL_ENABLED === 'true'),
      pollingInterval: ((ksm.deviceApproval?.pollingInterval || parseInt(process.env.DEVICE_POLL_INTERVAL)) || 120) * 1000,
    },

    // ==================== Feature Flags ====================
    features: {
      requireShareApproval: process.env.REQUIRE_SHARE_APPROVAL !== 'false',
      enableSearch: process.env.ENABLE_SEARCH !== 'false',
    },
    
    // ==================== KSM Status ====================
    ksm: {
      enabled: !!ksmData,
      loadedSections: ksmData ? Object.keys(ksmData) : [],
    },
  };
}


async function initializeConfig() {
  if (configInitialized) {
    return cachedConfig;
  }
  
  // Try to load from KSM
  const ksmData = await loadFromKsm();
  
  // Build and cache config
  cachedConfig = buildConfig(ksmData);
  configInitialized = true;
  
  // Log configuration source
  if (cachedConfig.ksm.enabled) {
    console.log(`[Config] Loaded from KSM: ${cachedConfig.ksm.loadedSections.join(', ')}`);
  } else {
    console.log('[Config] Loaded from environment variables');
  }
  
  return cachedConfig;
}


function getConfig() {
  if (cachedConfig) {
    return cachedConfig;
  }
  
  // Return default config from environment (for backwards compatibility)
  return buildConfig(null);
}

/**
 * Reset configuration (for testing)
 */
function resetConfig() {
  configInitialized = false;
  cachedConfig = null;
}


const defaultConfig = buildConfig(null);

module.exports = {
  // New async API
  initializeConfig,
  getConfig,
  resetConfig,
  
  // Backwards compatible: spread default config at module level
  ...defaultConfig,
};
