/**
 * KSM (Keeper Secrets Manager) Service
 * 
 * Fetches credentials from Keeper Secrets Manager for secure configuration.
 * Similar to the Slack app's ksm_utils.py implementation.
 */

const fs = require('fs');
const path = require('path');
const os = require('os');
const { createLogger } = require('./logger');

const log = createLogger('KsmService');

let SecretsManager = null;
let LocalConfigStorage = null;

/**
 * Check if KSM SDK is available
 */
function checkKsmDependency() {
  try {
    const ksm = require('@keeper-security/secrets-manager-core');
    SecretsManager = ksm.getSecrets;
    LocalConfigStorage = ksm.localConfigStorage;
    return true;
  } catch (error) {
    log.warn('KSM SDK not available:', error.message);
    return false;
  }
}

/**
 * Check if input is base64-encoded JSON
 */
function isBase64Config(inputStr) {
  if (!inputStr) return false;
  
  // If it looks like a file path, it's not base64
  if (inputStr.startsWith('/') || inputStr.startsWith('./') || 
      inputStr.startsWith('../') || inputStr.startsWith('~') ||
      fs.existsSync(inputStr)) {
    return false;
  }
  
  try {
    const decoded = Buffer.from(inputStr, 'base64').toString('utf-8');
    JSON.parse(decoded);
    return true;
  } catch {
    return false;
  }
}

/**
 * Process KSM config input - decode base64 or return file path
 */
function processKsmConfig(ksmConfigInput) {
  if (!ksmConfigInput) return null;
  
  if (isBase64Config(ksmConfigInput)) {
    try {
      const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ksm_'));
      const configPath = path.join(tmpDir, 'ksm-config.json');
      
      const decoded = Buffer.from(ksmConfigInput, 'base64').toString('utf-8');
      const configData = JSON.parse(decoded);
      
      if (typeof configData !== 'object') {
        log.error('Invalid KSM config format - must be JSON object');
        return null;
      }
      
      fs.writeFileSync(configPath, JSON.stringify(configData, null, 2), { mode: 0o600 });
      log.debug('Created temporary KSM config file');
      return configPath;
    } catch (error) {
      log.error('Failed to process KSM config:', error.message);
      return null;
    }
  } else {
    if (fs.existsSync(ksmConfigInput)) {
      return ksmConfigInput;
    } else {
      log.error('KSM config file not found:', ksmConfigInput);
      return null;
    }
  }
}

/**
 * Initialize SecretsManager with config
 */
async function initializeSecretsManager(ksmConfigPath) {
  if (!checkKsmDependency()) return null;
  
  try {
    const ksm = require('@keeper-security/secrets-manager-core');
    
    if (!fs.existsSync(ksmConfigPath)) {
      log.error('KSM config file not found');
      return null;
    }
    
    const storage = ksm.localConfigStorage(ksmConfigPath);
    return { ksm, storage };
  } catch (error) {
    log.error('Failed to initialize SecretsManager:', error.message);
    return null;
  }
}

/**
 * Get secret by UID or title
 */
async function getSecretByUidOrTitle(ksm, storage, recordIdentifier) {
  try {
    // First try to get by UID
    const secrets = await ksm.getSecrets({ storage }, [recordIdentifier]);
    if (secrets && secrets.records && secrets.records.length > 0) {
      return secrets.records[0];
    }
  } catch (error) {
    log.debug(`UID lookup failed for ${recordIdentifier}: ${error.message}, trying title lookup...`);
  }
  
  try {
    // Try to get all secrets and filter by title
    const allSecrets = await ksm.getSecrets({ storage });
    if (allSecrets && allSecrets.records) {
      const matchingRecords = allSecrets.records.filter(
        record => record.data && record.data.title === recordIdentifier
      );
      
      if (matchingRecords.length === 0) {
        log.error(`Record not found by UID or title: ${recordIdentifier}`);
        return null;
      } else if (matchingRecords.length > 1) {
        log.error(`Multiple records found with title '${recordIdentifier}' (${matchingRecords.length} records). Please use UID or a unique title.`);
        return null;
      } else {
        log.info(`Found record by title: ${recordIdentifier}`);
        return matchingRecords[0];
      }
    }
    return null;
  } catch (error) {
    log.error(`Failed to lookup record by title '${recordIdentifier}':`, error.message);
    return null;
  }
}

/**
 * Extract field value from KSM record
 */
function extractFieldValue(record, fieldLabel) {
  if (!record || !record.data) return null;
  
  const data = record.data;
  
  // Try fields array
  if (data.fields && Array.isArray(data.fields)) {
    for (const field of data.fields) {
      const label = field.label || field.type;
      if (label && label.toLowerCase() === fieldLabel.toLowerCase()) {
        const value = Array.isArray(field.value) ? field.value[0] : field.value;
        return value ? String(value).trim() : null;
      }
    }
  }
  
  // Try custom fields array
  if (data.custom && Array.isArray(data.custom)) {
    for (const field of data.custom) {
      const label = field.label || field.type;
      if (label && label.toLowerCase() === fieldLabel.toLowerCase()) {
        const value = Array.isArray(field.value) ? field.value[0] : field.value;
        return value ? String(value).trim() : null;
      }
    }
  }
  
  // Try with variations (underscore, hyphen)
  const variations = [
    fieldLabel.replace(/_/g, '-'),
    fieldLabel.replace(/-/g, '_'),
    fieldLabel.toLowerCase(),
    fieldLabel.toUpperCase(),
  ];
  
  for (const variation of variations) {
    if (variation === fieldLabel) continue;
    
    if (data.fields && Array.isArray(data.fields)) {
      for (const field of data.fields) {
        const label = field.label || field.type;
        if (label && label.toLowerCase() === variation.toLowerCase()) {
          const value = Array.isArray(field.value) ? field.value[0] : field.value;
          return value ? String(value).trim() : null;
        }
      }
    }
    
    if (data.custom && Array.isArray(data.custom)) {
      for (const field of data.custom) {
        const label = field.label || field.type;
        if (label && label.toLowerCase() === variation.toLowerCase()) {
          const value = Array.isArray(field.value) ? field.value[0] : field.value;
          return value ? String(value).trim() : null;
        }
      }
    }
  }
  
  return null;
}

/**
 * Extract notes field which may contain JSON config
 */
function extractNotesJson(record) {
  if (!record || !record.data) return null;
  
  const notes = extractFieldValue(record, 'notes');
  if (!notes) return null;
  
  try {
    return JSON.parse(notes);
  } catch {
    return null;
  }
}

/**
 * Fix service URL for Docker/production environment
 * 
 * This function is ONLY called for KSM-loaded URLs.
 * KSM is only used for Docker/production, so we always replace localhost.
 * Local development uses .localConfigs which has the correct localhost URLs.
 */
function fixServiceUrlForDocker(serviceUrl) {
  if (!serviceUrl) return serviceUrl;
  
  // Always replace localhost when called (only called for KSM-loaded URLs)
  // KSM is only used for Docker/production deployments
  if (serviceUrl.toLowerCase().includes('localhost') || serviceUrl.includes('127.0.0.1')) {
    log.info('Fixing service URL for Docker', { original: serviceUrl });
    // Replace localhost with commander-teams for Docker networking
    serviceUrl = serviceUrl.replace('localhost', 'commander-teams');
    serviceUrl = serviceUrl.replace('127.0.0.1', 'commander-teams');
    log.info('Fixed service URL for Docker', { fixed: serviceUrl });
  }
  
  return serviceUrl;
}

/**
 * Fetch credentials from KSM records
 * 
 * @param {string} ksmConfig - Base64 encoded KSM config or file path
 * @param {string} commanderRecordTitle - Title or UID of Commander config record
 * @param {string} teamsRecordTitle - Title or UID of Teams config record
 * @returns {Object} Configuration data from KSM
 */
async function fetchCredentialsFromKsm(ksmConfig, commanderRecordTitle, teamsRecordTitle) {
  const configData = {};
  
  if (!checkKsmDependency()) {
    log.warn('KSM not available, skipping KSM credential fetch');
    return configData;
  }
  
  // Process KSM config
  const ksmConfigPath = processKsmConfig(ksmConfig);
  if (!ksmConfigPath) {
    log.error('Failed to process KSM config');
    return configData;
  }
  
  // Initialize SecretsManager
  const sm = await initializeSecretsManager(ksmConfigPath);
  if (!sm) {
    log.error('Failed to initialize SecretsManager');
    return configData;
  }
  
  const { ksm, storage } = sm;
  
  // Fetch Commander/Keeper Service Mode credentials
  if (commanderRecordTitle) {
    try {
      const secret = await getSecretByUidOrTitle(ksm, storage, commanderRecordTitle);
      if (secret) {
        const keeperConfig = {};
        const notesJson = extractNotesJson(secret);
        
        // Extract fields (notes JSON takes precedence)
        let serviceUrl = notesJson?.service_url || 
                        extractFieldValue(secret, 'service_url') || 
                        extractFieldValue(secret, 'service-url');
        let apiKey = notesJson?.api_key || 
                    extractFieldValue(secret, 'api_key') || 
                    extractFieldValue(secret, 'api-key');
        
        if (serviceUrl) {
          keeperConfig.serviceUrl = fixServiceUrlForDocker(serviceUrl);
        }
        if (apiKey) {
          keeperConfig.apiKey = apiKey;
        }
        
        if (Object.keys(keeperConfig).length > 0) {
          configData.keeper = keeperConfig;
        } else {
          log.warn('No Keeper config extracted from KSM record');
        }
      }
    } catch (error) {
      log.error('Failed to fetch commander record:', error.message);
    }
  }
  
  // Fetch Teams Bot credentials
  if (teamsRecordTitle) {
    try {
      const secret = await getSecretByUidOrTitle(ksm, storage, teamsRecordTitle);
      if (secret) {
        const teamsConfig = {};
        const pedmConfig = {};
        const deviceApprovalConfig = {};
        const notesJson = extractNotesJson(secret);
        
        // Extract Teams/Azure fields
        const clientId = notesJson?.client_id || 
                        extractFieldValue(secret, 'client_id') || 
                        extractFieldValue(secret, 'client-id');
        const clientSecret = notesJson?.client_secret || 
                            extractFieldValue(secret, 'client_secret') || 
                            extractFieldValue(secret, 'client-secret') ||
                            extractFieldValue(secret, 'password');
        const tenantId = notesJson?.tenant_id || 
                        extractFieldValue(secret, 'tenant_id') || 
                        extractFieldValue(secret, 'tenant-id');
        const botType = notesJson?.bot_type || 
                       extractFieldValue(secret, 'bot_type') || 
                       extractFieldValue(secret, 'bot-type');
        const approvalsChannelId = notesJson?.approvals_channel_id || 
                                  extractFieldValue(secret, 'approvals_channel_id') || 
                                  extractFieldValue(secret, 'approvals-channel-id');
        const approvalsTeamId = notesJson?.approvals_team_id || 
                               extractFieldValue(secret, 'approvals_team_id') || 
                               extractFieldValue(secret, 'approvals-team-id');
        
        // Extract PEDM config
        const pedmEnabled = notesJson?.pedm_enabled ?? 
                           extractFieldValue(secret, 'pedm_enabled') ?? 
                           extractFieldValue(secret, 'pedm-enabled');
        const pedmInterval = notesJson?.pedm_polling_interval || 
                            extractFieldValue(secret, 'pedm_polling_interval') || 
                            extractFieldValue(secret, 'pedm-polling-interval');
        
        // Extract Device Approval config
        const deviceEnabled = notesJson?.device_approval_enabled ?? 
                             extractFieldValue(secret, 'device_approval_enabled') ?? 
                             extractFieldValue(secret, 'device-approval-enabled');
        const deviceInterval = notesJson?.device_approval_polling_interval || 
                              extractFieldValue(secret, 'device_approval_polling_interval') || 
                              extractFieldValue(secret, 'device-approval-polling-interval');
        
        // Build Teams config
        if (clientId) teamsConfig.clientId = clientId;
        if (clientSecret) teamsConfig.clientSecret = clientSecret;
        if (tenantId) teamsConfig.tenantId = tenantId;
        if (botType) teamsConfig.botType = botType;
        if (approvalsChannelId) teamsConfig.approvalsChannelId = approvalsChannelId;
        if (approvalsTeamId) teamsConfig.approvalsTeamId = approvalsTeamId;
        
        // Build PEDM config
        if (pedmEnabled !== null && pedmEnabled !== undefined) {
          pedmConfig.enabled = String(pedmEnabled).toLowerCase() === 'true' || pedmEnabled === true;
        }
        if (pedmInterval) {
          pedmConfig.pollingInterval = parseInt(pedmInterval, 10) || 120;
        }
        
        // Build Device Approval config
        if (deviceEnabled !== null && deviceEnabled !== undefined) {
          deviceApprovalConfig.enabled = String(deviceEnabled).toLowerCase() === 'true' || deviceEnabled === true;
        }
        if (deviceInterval) {
          deviceApprovalConfig.pollingInterval = parseInt(deviceInterval, 10) || 120;
        }
        
        // Validate required Teams fields
        const requiredFields = ['clientId', 'clientSecret', 'tenantId'];
        const missingFields = requiredFields.filter(f => !teamsConfig[f]);
        if (missingFields.length > 0) {
          log.warn(`Missing Teams fields in KSM record: ${missingFields.join(', ')}`);
        }
        
        if (Object.keys(teamsConfig).length > 0) {
          configData.teams = teamsConfig;
        } else {
          log.warn('No Teams config extracted from KSM record');
        }
        
        if (Object.keys(pedmConfig).length > 0) {
          configData.pedm = pedmConfig;
        }
        
        if (Object.keys(deviceApprovalConfig).length > 0) {
          configData.deviceApproval = deviceApprovalConfig;
        }
      }
    } catch (error) {
      log.error('Failed to fetch teams record:', error.message);
    }
  }
  
  // Summary message
  const fetchedItems = [];
  if (configData.keeper) fetchedItems.push('Service Mode Credentials');
  if (configData.teams) fetchedItems.push('Teams Credentials');
  
  if (fetchedItems.length > 0) {
    log.info(`Credentials fetched successfully from KSM vault: ${fetchedItems.join(', ')}`);
  }
  
  return configData;
}

module.exports = {
  checkKsmDependency,
  fetchCredentialsFromKsm,
  processKsmConfig,
  extractFieldValue,
  fixServiceUrlForDocker,
};
