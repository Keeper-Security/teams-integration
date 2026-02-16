/**
 * Device Approval Poller
 * 
 * Background service that periodically polls for pending Cloud SSO
 * device approval requests and posts them to the approvals channel.
 */

const keeperClient = require('../services/keeperClient');
const { getChannelService } = require('../services');
const cards = require('../cards');
const config = require('../config');

class DevicePoller {
  constructor(teamsApp) {
    this.teamsApp = teamsApp;
    this.interval = config.deviceApproval?.pollingInterval || 120000;
    this.enabled = config.deviceApproval?.enabled || false;
    this.timer = null;
    this.processedDevices = new Map(); // Track already-posted devices with timestamps
    this.consecutiveErrors = 0;
    this.maxErrors = 3;
  }

  /**
   * Start the polling service
   */
  start() {
    if (!this.enabled) {
      console.log('[Device Poller] Disabled in config');
      return;
    }

    console.log('[Device Poller] Starting with interval: ' + this.interval + 'ms');
    
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
      console.log('[Device Poller] Stopped');
    }
  }

  /**
   * Poll for pending device approval requests
   */
  async poll() {
    console.log('[Device Poller] Polling for pending device approvals...');
    
    try {
      const devices = await keeperClient.getPendingDeviceApprovals();
      
      if (!devices || devices.length === 0) {
        console.log('[Device Poller] No pending device approvals');
        return;
      }
      
      console.log('[Device Poller] Found ' + devices.length + ' pending devices');
      
      for (const device of devices) {
        const deviceId = device.device_id || device.deviceId || device.id;
        
        if (!deviceId) {
          console.log('[Device Poller] Skipping device without ID');
          continue;
        }
        
        if (this.processedDevices.has(deviceId)) {
          console.log('[Device Poller] Already processed: ' + deviceId);
          continue;
        }
        
        await this.postApprovalCard(device);
        this.processedDevices.set(deviceId, Date.now());
      }
    } catch (error) {
      console.error('[Device Poller] Error polling:', error.message);
    }
  }

  /**
   * Post a device approval card to the approvals channel
   */
  async postApprovalCard(device) {
    const deviceId = device.device_id || device.deviceId || device.id;
    
    console.log('[Device Poller] Posting card for: ' + deviceId);
    
    const card = cards.buildDeviceApprovalCard({
      deviceId: deviceId,
      deviceName: device.device_name || device.deviceName || device.name,
      deviceType: device.device_type || device.deviceType || device.type,
      username: device.username || device.user,
      email: device.email || device.user_email,
      ipAddress: device.ip_address || device.ipAddress || device.ip,
      location: device.location,
      created: device.created || device.timestamp,
    });
    
    // Use ChannelService to send to approvals channel
    const channelService = getChannelService();
    
    if (channelService && channelService.isApprovalsChannelReady()) {
      const deviceName = device.device_name || device.deviceName || device.name || 'Unknown Device';
      const userEmail = device.email || device.user_email || 'Unknown User';
      
      const sent = await channelService.sendApprovalCard(
        card,
        `**Device Approval Request** - ${deviceName} (${userEmail})`
      );
      
      if (sent) {
        console.log('[Device Poller] Card posted to approvals channel: ' + deviceId);
        this.consecutiveErrors = 0;
        return true;
      } else {
        console.warn('[Device Poller] Failed to post card to channel: ' + deviceId);
        this.consecutiveErrors++;
        return false;
      }
    } else {
      console.warn('[Device Poller] Approvals channel not ready - card not sent');
      console.warn('[Device Poller] Send a message in the approvals channel to initialize');
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
        cleaned++;
      }
    }
    
    if (cleaned > 0) {
      console.log('[Device Poller] Cleaned up ' + cleaned + ' old entries');
    }
  }
}

module.exports = DevicePoller;
