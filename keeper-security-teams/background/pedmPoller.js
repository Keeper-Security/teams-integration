/**
 * PEDM (Privileged Elevation & Delegation Management) Poller
 * 
 * Background service that periodically polls for pending PEDM requests
 * and posts them to the approvals channel in Teams.
 */

const keeperClient = require('../services/keeperClient');
const { getChannelService, storeApprovalActivityId, createLogger } = require('../services');
const cards = require('../cards');
const { getConfig } = require('../config');

const log = createLogger('PedmPoller');


function parsePedmRequest(data) {
  let username = '';
  const accountInfo = data.account_info || [];
  for (const info of accountInfo) {
    if (typeof info === 'string' && info.startsWith('Username=')) {
      username = info.split('=')[1] || '';
      break;
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
    
    if (info.startsWith('Description=')) {
      description = info.split('=').slice(1).join('=') || '';
    } else if (info.startsWith('FileName=')) {
      fileName = info.split('=').slice(1).join('=') || '';
    } else if (info.startsWith('FilePath=')) {
      filePath = info.split('=').slice(1).join('=') || '';
    } else if (info.startsWith('CommandLine=')) {
      command = info.split('=').slice(1).join('=') || '';
    }
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

class PedmPoller {
  constructor(teamsApp) {
    this.teamsApp = teamsApp;
    const config = getConfig();
    this.interval = config.pedm?.pollingInterval || 120000;
    this.enabled = config.pedm?.enabled || false;
    this.timer = null;
    this.processedRequests = new Map(); // Track already-posted requests with timestamps
    this.consecutiveErrors = 0;
    this.maxErrors = 3;
  }

  /**
   * Start the polling service
   */
  start() {
    if (!this.enabled) {
      log.info('Disabled in config');
      return;
    }

    log.info(`Starting with interval: ${this.interval}ms`);
    
    // Run immediately on start
    this.poll();
    
    // Then run on interval
    this.timer = setInterval(() => this.poll(), this.interval);
    
    // Cleanup old entries every hour
    setInterval(() => this.cleanup(), 3600000);
  }

  /**
   * Stop the polling service
   */
  stop() {
    if (this.timer) {
      clearInterval(this.timer);
      this.timer = null;
      log.info('Stopped');
    }
  }

  /**
   * Poll for pending PEDM requests
   */
  async poll() {
    log.debug('Polling for pending requests...');
    
    try {
      const requests = await keeperClient.getPendingPedmRequests();
      
      if (!requests || requests.length === 0) {
        log.debug('No pending requests');
        return;
      }
      
      log.debug(`Found ${requests.length} pending requests`);
      
      for (const request of requests) {
        const requestId = request.approval_uid || request.approvalUid || request.id;
        
        if (!requestId) {
          log.debug('Skipping request without ID');
          continue;
        }
        
        if (this.processedRequests.has(requestId)) {
          log.debug('Already processed', requestId);
          continue;
        }
        
        await this.postApprovalCard(request);
        this.processedRequests.set(requestId, Date.now());
      }
    } catch (error) {
      log.error('Error polling', error.message);
    }
  }

  /**
   * Post a PEDM approval card to the approvals channel
   */
  async postApprovalCard(rawRequest) {
    // Parse the request to extract fields from arrays
    const request = parsePedmRequest(rawRequest);
    
    // Use approval_uid as the approvalId for activity tracking
    const approvalId = 'epm_' + (request.approvalUid || Math.random().toString(36).substring(2, 10));
    
    log.debug('Posting card for', { approvalUid: request.approvalUid, approvalId });
    log.debug('Parsed request', { approvalUid: request.approvalUid, username: request.username, agentUid: request.agentUid });
    
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
        log.info('Card posted to approvals channel', { approvalUid: request.approvalUid, approvalId, activityId: result.activityId });
        this.consecutiveErrors = 0;
        return true;
      } else {
        log.warn('Failed to post card to channel', { approvalUid: request.approvalUid });
        this.consecutiveErrors++;
        return false;
      }
    } else {
      log.warn('Approvals channel not ready - card not sent');
      return false;
    }
  }

  /**
   * Clean up old processed requests (to prevent memory leak)
   */
  cleanup(maxAge = 86400000) { // Default: 24 hours
    const now = Date.now();
    let cleaned = 0;
    
    for (const [id, timestamp] of this.processedRequests) {
      if (now - timestamp > maxAge) {
        this.processedRequests.delete(id);
        cleaned++;
      }
    }
    
    if (cleaned > 0) {
      log.debug(`Cleaned up ${cleaned} old entries`);
    }
  }
}

module.exports = PedmPoller;
