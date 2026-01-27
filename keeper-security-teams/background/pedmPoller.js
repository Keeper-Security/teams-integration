/**
 * PEDM (Privileged Elevation & Delegation Management) Poller
 * 
 * Background service that periodically polls for pending PEDM requests
 * and posts them to the approvals channel in Teams.
 */

const keeperClient = require('../services/keeperClient');
const { getChannelService } = require('../services');
const cards = require('../cards');
const config = require('../config');

class PedmPoller {
  constructor(teamsApp) {
    this.teamsApp = teamsApp;
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
      console.log('[PEDM Poller] Disabled in config');
      return;
    }

    console.log('[PEDM Poller] Starting with interval: ' + this.interval + 'ms');
    
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
      console.log('[PEDM Poller] Stopped');
    }
  }

  /**
   * Poll for pending PEDM requests
   */
  async poll() {
    console.log('[PEDM Poller] Polling for pending requests...');
    
    try {
      const requests = await keeperClient.getPendingPedmRequests();
      
      if (!requests || requests.length === 0) {
        console.log('[PEDM Poller] No pending requests');
        return;
      }
      
      console.log('[PEDM Poller] Found ' + requests.length + ' pending requests');
      
      for (const request of requests) {
        const requestId = request.approval_uid || request.approvalUid || request.id;
        
        if (!requestId) {
          console.log('[PEDM Poller] Skipping request without ID');
          continue;
        }
        
        if (this.processedRequests.has(requestId)) {
          console.log('[PEDM Poller] Already processed: ' + requestId);
          continue;
        }
        
        await this.postApprovalCard(request);
        this.processedRequests.set(requestId, Date.now());
      }
    } catch (error) {
      console.error('[PEDM Poller] Error polling:', error.message);
    }
  }

  /**
   * Post a PEDM approval card to the approvals channel
   */
  async postApprovalCard(request) {
    const approvalUid = request.approval_uid || request.approvalUid || request.id;
    
    console.log('[PEDM Poller] Posting card for: ' + approvalUid);
    
    const card = cards.buildPedmApprovalCard({
      approvalUid: approvalUid,
      approvalType: request.approval_type || request.type || 'PrivilegeElevation',
      username: request.username || request.user || 'Unknown',
      command: request.command_line || request.command,
      fileName: request.file_name || request.fileName,
      filePath: request.file_path || request.filePath,
      description: request.description,
      justification: request.justification,
      expireIn: request.expire_in || request.expireIn,
      created: request.created || request.timestamp,
    });
    
    // Use ChannelService to send to approvals channel
    const channelService = getChannelService();
    
    if (channelService && channelService.isApprovalsChannelReady()) {
      const sent = await channelService.sendApprovalCard(
        card,
        `⚡ **Privilege Elevation Request** from ${request.username || 'Unknown User'}`
      );
      
      if (sent) {
        console.log('[PEDM Poller] Card posted to approvals channel: ' + approvalUid);
        this.consecutiveErrors = 0;
        return true;
      } else {
        console.warn('[PEDM Poller] Failed to post card to channel: ' + approvalUid);
        this.consecutiveErrors++;
        return false;
      }
    } else {
      console.warn('[PEDM Poller] Approvals channel not ready - card not sent');
      console.warn('[PEDM Poller] Send a message in the approvals channel to initialize');
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
      console.log('[PEDM Poller] Cleaned up ' + cleaned + ' old entries');
    }
  }
}

module.exports = PedmPoller;
