/**
 * Device Approval Poller
 * 
 * Background service that periodically polls for pending Cloud SSO
 * device approval requests and posts them to the approvals channel.
 */

const keeperClient = require('../services/keeperClient');
const { getChannelService, storeApprovalActivityId, createLogger } = require('../services');
const cards = require('../cards');
const { getConfig } = require('../config');

const log = createLogger('DevicePoller');

class DevicePoller {
  constructor(teamsApp) {
    this.teamsApp = teamsApp;
    const config = getConfig();
    this.interval = config.deviceApproval?.pollingInterval || 120000;
    this.enabled = config.deviceApproval?.enabled || false;
    this.timer = null;
    this.processedDevices = new Map(); 
    this.deviceToApprovalId = new Map();
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
   * Poll for pending device approval requests
   */
  async poll() {
    log.debug('Polling for pending device approvals...');
    
    try {
      const devices = await keeperClient.getPendingDeviceApprovals();
      
      if (!devices || devices.length === 0) {
        log.debug('No pending device approvals');
        return;
      }
      
      log.debug(`Found ${devices.length} pending devices`);
      
      for (const device of devices) {
        const deviceId = device.device_id || device.deviceId || device.id;
        
        if (!deviceId) {
          log.debug('Skipping device without ID');
          continue;
        }
        
        if (this.processedDevices.has(deviceId)) {
          log.debug('Already processed', deviceId);
          continue;
        }
        
        await this.postApprovalCard(device);
        this.processedDevices.set(deviceId, Date.now());
      }
    } catch (error) {
      log.error('Error polling', error.message);
    }
  }

  /**
   * Generate a unique approval ID for tracking
   */
  generateApprovalId() {
    return 'dev_' + Math.random().toString(36).substring(2, 10);
  }

  /**
   * Post a device approval card to the approvals channel
   */
  async postApprovalCard(device) {
    const deviceId = device.device_id || device.deviceId || device.id;
    
    log.debug('Posting card for device', { deviceId });
    
    // Generate a unique approval ID for this device card
    const approvalId = this.generateApprovalId();
    
    const card = cards.buildDeviceApprovalCard({
      deviceId: deviceId,
      deviceName: device.device_name || device.deviceName || device.name,
      deviceType: device.device_type || device.deviceType || device.type,
      username: device.username || device.user,
      email: device.email || device.user_email,
      ipAddress: device.ip_address || device.ipAddress || device.ip,
      location: device.location,
      created: device.created || device.timestamp,
      approvalId: approvalId, // Include approvalId in card data for tracking
    });
    
    // Use ChannelService to send to approvals channel
    const channelService = getChannelService();
    
    if (channelService && channelService.isApprovalsChannelReady()) {
      const deviceName = device.device_name || device.deviceName || device.name || 'Unknown Device';
      const userEmail = device.email || device.user_email || 'Unknown User';
      
      // Use sendApprovalCardViaConnector to get proper activity ID
      const result = await channelService.sendApprovalCardViaConnector(
        card,
        approvalId,
        `**Device Approval Request** - ${deviceName} (${userEmail})`
      );
      
      if (result.success) {
        log.info('Card posted to approvals channel', { deviceId, approvalId, activityId: result.activityId });
        this.deviceToApprovalId.set(deviceId, approvalId);
        this.consecutiveErrors = 0;
        return true;
      } else {
        log.warn('Failed to post card to channel', { deviceId });
        this.consecutiveErrors++;
        return false;
      }
    } else {
      log.warn('Approvals channel not ready - card not sent');
      return false;
    }
  }

  /**
   * Clean up old processed devices (to prevent memory leak)
   */
  cleanup(maxAge = 86400000) { // Default: 24 hours
    const now = Date.now();
    let cleaned = 0;
    
    for (const [id, timestamp] of this.processedDevices) {
      if (now - timestamp > maxAge) {
        this.processedDevices.delete(id);
        this.deviceToApprovalId.delete(id); // Also clean up approval ID mapping
        cleaned++;
      }
    }
    
    if (cleaned > 0) {
      log.debug(`Cleaned up ${cleaned} old entries`);
    }
  }
}

module.exports = DevicePoller;
