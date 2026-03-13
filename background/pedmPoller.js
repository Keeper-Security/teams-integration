/**
 * EPM (Endpoint Privilege Manager) Poller
 * 
 * Background service that periodically polls for pending EPM requests
 * and posts them to the approvals channel in Teams.
 */

const keeperClient = require('../services/keeperClient');
const { getChannelService, storeApprovalActivityId, createLogger } = require('../services');
const cards = require('../cards');
const { getConfig } = require('../config');

const log = createLogger('EpmPoller');

/**
 * Extract value from a key=value string if it matches the prefix
 * Handles values that may contain '=' characters
 * @param {string} str - The string to parse (e.g., "FilePath=C:\Program Files\app.exe")
 * @param {string} prefix - The prefix to match (e.g., "FilePath=")
 * @returns {string|null} The extracted value or null if prefix doesn't match
 */
function extractValue(str, prefix) {
  if (str.startsWith(prefix)) {
    return str.slice(prefix.length) || '';
  }
  return null;
}

/**
 * Parse EPM request data from API response
 */
function parseEpmRequest(data) {
  let username = '';
  const accountInfo = data.account_info || [];
  for (const info of accountInfo) {
    if (typeof info === 'string') {
      const value = extractValue(info, 'Username=');
      if (value !== null) {
        username = value;
        break;
      }
    }
  }
  
  // Extract fields from application_info array
  let description = '';
  let fileName = '';
  let filePath = '';
  let command = '';
  
  const applicationInfo = data.application_info || [];
  for (const info of applicationInfo) {
    if (typeof info !== 'string') continue;
    
    description = extractValue(info, 'Description=') ?? description;
    fileName = extractValue(info, 'FileName=') ?? fileName;
    filePath = extractValue(info, 'FilePath=') ?? filePath;
    command = extractValue(info, 'CommandLine=') ?? command;
  }
  
  // Return parsed request with fallbacks to direct properties
  return {
    approvalUid: data.approval_uid || data.approvalUid || data.id || '',
    approvalType: data.approval_type || data.type || 'PrivilegeElevation',
    status: data.status || 'Pending',
    agentUid: data.agent_uid || data.agentUid || '',
    username: username || data.username || data.user || 'Unknown',
    command: command || data.command_line || data.command || '',
    fileName: fileName || data.file_name || data.fileName || '',
    filePath: filePath || data.file_path || data.filePath || '',
    description: description || data.description || '',
    justification: data.justification || '',
    expireIn: data.expire_in || data.expireIn || 30,
    created: data.created || data.timestamp || '',
  };
}

class EpmPoller {
  constructor(teamsApp) {
    this.teamsApp = teamsApp;
    const config = getConfig();
    this.interval = config.pedm?.pollingInterval || 120000;
    this.enabled = config.pedm?.enabled || false;
    this.timer = null;
    this.seenRequestIds = new Set(); // Track seen request IDs (cleared when no longer pending)
    this.consecutiveErrors = 0;
    this.maxErrors = 3;
  }

  /**
   * Start the EPM polling service
   */
  start() {
    if (!this.enabled) {
      log.info('EPM polling is disabled in config');
      return;
    }

    log.info(`EPM poller starting with interval: ${this.interval}ms`);
    
    // Run immediately on start
    this.poll();
    
    // Then run on interval
    this.timer = setInterval(() => this.poll(), this.interval);
  }

  /**
   * Stop the EPM polling service
   */
  stop() {
    if (this.timer) {
      clearInterval(this.timer);
      this.timer = null;
      log.info('EPM poller stopped');
    }
  }

  /**
   * Poll for pending EPM requests
   * Uses Slack-style approach: relies on Keeper API as source of truth.
   * Once a request is approved/denied, it won't appear in pending list.
   */
  async poll() {
    log.info('Polling for pending EPM requests...');
    
    try {
      const requests = await keeperClient.getPendingPedmRequests();
      
      // null/undefined means API failure - keep seen list intact
      if (requests === null || requests === undefined) {
        log.debug('EPM API failed/timed out, keeping seen list intact');
        return;
      }
      
      // No pending requests - clear the seen list
      if (requests.length === 0) {
        if (this.seenRequestIds.size > 0) {
          log.debug('No pending EPM requests, clearing seen list');
          this.seenRequestIds.clear();
        }
        log.info('No pending EPM requests found');
        return;
      }
      
      log.info(`Found ${requests.length} pending EPM requests`);
      
      // Track current pending IDs and identify new requests
      const currentIds = new Set();
      const newRequests = [];
      
      for (const request of requests) {
        const requestId = request.approval_uid || request.approvalUid || request.id;
        
        if (!requestId) {
          log.debug('Skipping EPM request without ID');
          continue;
        }
        
        currentIds.add(requestId);
        
        // Check if this is a NEW request
        if (!this.seenRequestIds.has(requestId)) {
          newRequests.push(request);
          this.seenRequestIds.add(requestId);
          log.info('New EPM request detected', { requestId });
        }
      }
      
      // Post only NEW requests to Teams
      if (newRequests.length > 0) {
        log.info(`Posting ${newRequests.length} new EPM request(s) to Teams`);
        for (const request of newRequests) {
          await this.postApprovalCard(request);
        }
      }
      
      // Cleanup: remove IDs that are no longer pending (processed elsewhere)
      const removedIds = [...this.seenRequestIds].filter(id => !currentIds.has(id));
      if (removedIds.length > 0) {
        log.debug(`Cleaning up ${removedIds.length} resolved EPM request(s)`);
        for (const id of removedIds) {
          this.seenRequestIds.delete(id);
        }
      }
      
      this.consecutiveErrors = 0;
    } catch (error) {
      this.consecutiveErrors++;
      log.error(`Error occurred while EPM polling (${this.consecutiveErrors}/${this.maxErrors})`, error.message);
      
      if (this.consecutiveErrors >= this.maxErrors) {
        log.warn('EPM polling stopped due to consecutive errors (feature may not be available)');
        this.stop();
      }
    }
  }

  /**
   * Post an EPM approval card to the approvals channel
   */
  async postApprovalCard(rawRequest) {
    // Parse the request to extract fields from arrays
    const request = parseEpmRequest(rawRequest);
    
    // Use approval_uid as the approvalId for activity tracking
    const approvalId = 'epm_' + (request.approvalUid || Math.random().toString(36).substring(2, 10));
    
    log.info('Posting EPM approval card', { approvalUid: request.approvalUid, approvalId });
    log.debug('Parsed EPM request', { approvalUid: request.approvalUid, username: request.username, agentUid: request.agentUid });
    
    const card = cards.buildPedmApprovalCard({
      approvalUid: request.approvalUid,
      approvalType: request.approvalType,
      agentUid: request.agentUid,
      username: request.username,
      command: request.command,
      fileName: request.fileName,
      filePath: request.filePath,
      description: request.description,
      justification: request.justification,
      expireIn: request.expireIn,
      created: request.created,
    });
    
    // Use ChannelService to send to approvals channel
    const channelService = getChannelService();
    
    if (channelService && channelService.isApprovalsChannelReady()) {
      // Use sendApprovalCardViaConnector to get proper activity ID
      const result = await channelService.sendApprovalCardViaConnector(
        card,
        approvalId,
        `**Endpoint Privilege Manager Request** from ${request.username}`
      );
      
      if (result.success) {
        log.info('EPM card posted to approvals channel', { approvalUid: request.approvalUid, approvalId, activityId: result.activityId });
        this.consecutiveErrors = 0;
        return true;
      } else {
        log.warn('Failed to post EPM card to channel', { approvalUid: request.approvalUid });
        this.consecutiveErrors++;
        return false;
      }
    } else {
      log.warn('Approvals channel not ready - EPM card not sent');
      return false;
    }
  }
}

module.exports = EpmPoller;
