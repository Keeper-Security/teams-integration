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
    this.seenDeviceIds = new Set(); // Track seen device IDs (cleared when no longer pending)
    this.deviceToApprovalId = new Map();
    this.consecutiveErrors = 0;
    this.maxErrors = 3;
  }

  /**
   * Start the device polling service
   */
  start() {
    if (!this.enabled) {
      log.info('Device poller is disabled in config');
      return;
    }

    log.info(`Device poller starting with interval: ${this.interval}ms`);
    
    // Run immediately on start
    this.poll();
    
    // Then run on interval
    this.timer = setInterval(() => this.poll(), this.interval);
  }

  /**
   * Stop the device polling service
   */
  stop() {
    if (this.timer) {
      clearInterval(this.timer);
      this.timer = null;
      log.info('Device poller stopped');
    }
  }

  /**
   * Poll for pending device approval requests
   * Uses Slack-style approach: relies on Keeper API as source of truth.
   * Once a device is approved/denied, it won't appear in pending list.
   */
  async poll() {
    log.info('Polling for pending device approvals...');
    
    try {
      const devices = await keeperClient.getPendingDeviceApprovals();
      
      // null/undefined means API failure - keep seen list intact
      if (devices === null || devices === undefined) {
        log.debug('Device API failed/timed out, keeping seen list intact');
        return;
      }
      
      // No pending devices - clear the seen list
      if (devices.length === 0) {
        if (this.seenDeviceIds.size > 0) {
          log.debug('No pending device approvals, clearing seen list');
          this.seenDeviceIds.clear();
          this.deviceToApprovalId.clear();
        }
        log.info('No pending device approvals found');
        return;
      }
      
      log.info(`Found ${devices.length} pending device approvals`);
      
      // Track current pending IDs and identify new devices
      const currentIds = new Set();
      const newDevices = [];
      
      for (const device of devices) {
        const deviceId = device.device_id || device.deviceId || device.id;
        
        if (!deviceId) {
          log.debug('Skipping device without ID');
          continue;
        }
        
        currentIds.add(deviceId);
        
        // Check if this is a NEW device
        if (!this.seenDeviceIds.has(deviceId)) {
          newDevices.push(device);
          this.seenDeviceIds.add(deviceId);
          const deviceName = device.device_name || device.deviceName || device.name || 'Unknown';
          log.info('New device approval request detected', { deviceId, deviceName });
        }
      }
      
      // Post only NEW devices to Teams
      if (newDevices.length > 0) {
        log.info(`Posting ${newDevices.length} new device approval(s) to Teams`);
        for (const device of newDevices) {
          await this.postApprovalCard(device);
        }
      }
      
      // Cleanup: remove IDs that are no longer pending (processed elsewhere)
      const removedIds = [...this.seenDeviceIds].filter(id => !currentIds.has(id));
      if (removedIds.length > 0) {
        log.debug(`Cleaning up ${removedIds.length} resolved device approval(s)`);
        for (const id of removedIds) {
          this.seenDeviceIds.delete(id);
          this.deviceToApprovalId.delete(id);
        }
      }
      
      this.consecutiveErrors = 0;
    } catch (error) {
      this.consecutiveErrors++;
      log.error(`Error occurred while device polling (${this.consecutiveErrors}/${this.maxErrors})`, error.message);
      
      if (this.consecutiveErrors >= this.maxErrors) {
        log.warn('Device polling stopped due to consecutive errors (feature may not be available)');
        this.stop();
      }
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
    
    log.info('Posting device approval card', { deviceId });
    
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
        log.info('Device approval card posted to channel', { deviceId, approvalId, activityId: result.activityId });
        this.deviceToApprovalId.set(deviceId, approvalId);
        this.consecutiveErrors = 0;
        return true;
      } else {
        log.warn('Failed to post device approval card to channel', { deviceId });
        this.consecutiveErrors++;
        return false;
      }
    } else {
      log.warn('Approvals channel not ready - device card not sent');
      return false;
    }
  }
}

module.exports = DevicePoller;
