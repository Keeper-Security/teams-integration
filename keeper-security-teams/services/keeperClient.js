/**
 * Keeper Commander Service Mode API Client
 * 
 * This module handles all communication with Keeper Commander running in Service Mode.
 * It provides methods for searching records/folders, granting access, creating shares, etc.
 */

const axios = require('axios');
const config = require('../config');

/**
 * Shell escape a string for safe command execution
 * @param {string} str - String to escape
 * @returns {string} - Escaped string
 */
function shellEscape(str) {
  if (!str) return '""';
  if (/[\s"'\\]/.test(str)) {
    return '"' + str.replace(/["\\]/g, '\\$&') + '"';
  }
  return str;
}

class KeeperClient {
  constructor() {
    this.baseUrl = (config.keeper?.serviceUrl || 'http://localhost:3001/api/v2/').replace(/\/$/, '');
    this.apiKey = config.keeper?.apiKey;
    
    this.client = axios.create({
      baseURL: this.baseUrl,
      timeout: 30000,
      headers: {
        'Content-Type': 'application/json',
        ...(this.apiKey && { 'api-key': this.apiKey }),
      },
    });
  }

  updateCredentials(serviceUrl, apiKey = null) {
    this.baseUrl = serviceUrl.replace(/\/$/, '');
    this.apiKey = apiKey;
    
    this.client = axios.create({
      baseURL: this.baseUrl,
      timeout: 30000,
      headers: {
        'Content-Type': 'application/json',
        ...(apiKey && { 'api-key': apiKey }),
      },
    });
    
    console.log('[KeeperClient] Credentials updated:', this.baseUrl);
  }

  // ==================== Health & Connection ====================

  async healthCheck() {
    try {
      const response = await this.client.get('/queue/status', { timeout: 5000 });
      return response.status === 200;
    } catch (error) {
      console.error('[KeeperClient] Health check failed:', error.message);
      return false;
    }
  }

  // ==================== Search Operations ====================

  async searchRecords(query, limit = 20) {
    try {
      const command = 'search -c r ' + shellEscape(query) + ' --format=json';
      const result = await this._executeCommandAsync(command, 30);
      
      if (!result || result.status !== 'success') {
        console.log('[KeeperClient] Search records failed:', result?.message);
        return [];
      }
      
      return this._parseSearchResults(result.data, 'record', limit);
    } catch (error) {
      console.error('[KeeperClient] Error searching records:', error.message);
      return [];
    }
  }

  async searchFolders(query, limit = 20) {
    try {
      const command = 'search -c s ' + shellEscape(query) + ' --format=json';
      const result = await this._executeCommandAsync(command, 30);
      
      if (!result || result.status !== 'success') {
        console.log('[KeeperClient] Search folders failed:', result?.message);
        return [];
      }
      
      return this._parseSearchResults(result.data, 'folder', limit);
    } catch (error) {
      console.error('[KeeperClient] Error searching folders:', error.message);
      return [];
    }
  }

  async getRecordByUid(recordUid) {
    try {
      const command = 'search ' + shellEscape(recordUid) + ' --format=json';
      const result = await this._executeCommandAsync(command, 15);
      
      if (!result || result.status !== 'success') {
        console.log('[KeeperClient] Get record failed:', result?.message);
        return null;
      }
      
      const data = result.data;
      if (!data || !Array.isArray(data) || data.length === 0) {
        return null;
      }
      
      const item = data[0];
      const itemType = item.type || 'record';
      
      let recordType = 'login';
      let notes = '';
      
      if (item.details) {
        const parts = item.details.split(', ');
        for (const part of parts) {
          if (part.startsWith('Type: ')) {
            recordType = part.replace('Type: ', '').trim();
          } else if (part.startsWith('Description: ')) {
            notes = part.replace('Description: ', '').trim();
          }
        }
      }
      
      return {
        uid: item.uid || recordUid,
        title: item.name || 'Untitled Record',
        recordType: itemType.includes('folder') ? itemType : recordType,
        folderUid: null,
        notes: notes,
      };
    } catch (error) {
      console.error('[KeeperClient] Error getting record:', error.message);
      return null;
    }
  }

  async getFolderByUid(folderUid) {
    try {
      const command = 'search ' + shellEscape(folderUid) + ' --format=json';
      const result = await this._executeCommandAsync(command, 15);
      
      if (!result || result.status !== 'success') {
        console.log('[KeeperClient] Get folder failed:', result?.message);
        return null;
      }
      
      const data = result.data;
      if (!data || !Array.isArray(data) || data.length === 0) {
        return null;
      }
      
      const item = data[0];
      
      return {
        uid: item.uid || folderUid,
        name: item.name || 'Untitled Folder',
        parentUid: null,
        folderType: item.type || 'shared_folder',
      };
    } catch (error) {
      console.error('[KeeperClient] Error getting folder:', error.message);
      return null;
    }
  }

  // ==================== Access Grant Operations ====================

  async grantRecordAccess(recordUid, userEmail, permission, durationSeconds = 86400) {
    try {
      if (permission === 'change_owner') {
        const command = 'share-record ' + recordUid + ' -e ' + userEmail + ' -a owner --force';
        const result = await this._executeCommandAsync(command, 15);
        
        if (result?.status === 'success') {
          return {
            success: true,
            expiresAt: 'N/A (Ownership Transfer)',
            permission: permission,
            duration: 'permanent',
          };
        } else {
          return { success: false, error: result?.message || 'Failed to transfer ownership' };
        }
      }
      
      const permissionFlags = [];
      if (permission === 'can_edit' || permission === 'edit_and_share') {
        permissionFlags.push('-w');
      }
      if (permission === 'can_share' || permission === 'edit_and_share') {
        permissionFlags.push('-s');
      }
      
      const permanentOnlyPermissions = ['can_share', 'edit_and_share', 'change_owner'];
      const isPermanent = permanentOnlyPermissions.includes(permission) || durationSeconds === null;
      
      try {
        const revokeCmd = 'share-record ' + recordUid + ' -e ' + userEmail + ' -a revoke --force';
        await this._executeCommandAsync(revokeCmd, 5);
      } catch (e) {
        // Ignore revoke errors
      }
      
      let command = 'share-record ' + recordUid + ' -e ' + userEmail + ' -a grant';
      if (permissionFlags.length > 0) {
        command += ' ' + permissionFlags.join(' ');
      }
      
      let expiresAtStr = 'Never (Permanent)';
      if (!isPermanent && durationSeconds) {
        const expireIn = this._formatDuration(durationSeconds);
        command += ' --expire-in ' + expireIn;
        const expiresAt = new Date(Date.now() + durationSeconds * 1000);
        expiresAtStr = expiresAt.toISOString();
      }
      
      command += ' --force';
      
      const result = await this._executeCommandAsync(command, 15);
      
      if (result?.status === 'success') {
        return {
          success: true,
          expiresAt: expiresAtStr,
          permission: permission,
          duration: isPermanent ? 'permanent' : 'temporary',
        };
      } else {
        const errorMsg = this._formatError(result?.message);
        return { success: false, error: errorMsg };
      }
    } catch (error) {
      return { success: false, error: 'Error granting record access: ' + error.message };
    }
  }

  async grantFolderAccess(folderUid, userEmail, permission, durationSeconds = 86400) {
    try {
      const permissionFlags = [];
      
      switch (permission) {
        case 'no_permissions':
          permissionFlags.push('-o', 'off', '-p', 'off');
          break;
        case 'manage_users':
          permissionFlags.push('-o', 'on', '-p', 'off');
          break;
        case 'manage_records':
          permissionFlags.push('-o', 'off', '-p', 'on');
          break;
        case 'manage_all':
          permissionFlags.push('-o', 'on', '-p', 'on');
          break;
      }
      
      const permanentOnlyPermissions = ['manage_users', 'manage_all'];
      const isPermanent = permanentOnlyPermissions.includes(permission) || durationSeconds === null;
      
      let command = 'share-folder ' + folderUid + ' -e ' + userEmail + ' -a grant';
      if (permissionFlags.length > 0) {
        command += ' ' + permissionFlags.join(' ');
      }
      
      let expiresAtStr = 'Never (Permanent)';
      if (!isPermanent && durationSeconds) {
        const expireIn = this._formatDuration(durationSeconds);
        command += ' --expire-in ' + expireIn;
        const expiresAt = new Date(Date.now() + durationSeconds * 1000);
        expiresAtStr = expiresAt.toISOString();
      }
      
      command += ' -f';
      
      const result = await this._executeCommandAsync(command, 15);
      
      if (result?.status === 'success') {
        return {
          success: true,
          expiresAt: expiresAtStr,
          permission: permission,
          duration: isPermanent ? 'permanent' : 'temporary',
        };
      } else {
        const errorMsg = this._formatError(result?.message);
        return { success: false, error: errorMsg };
      }
    } catch (error) {
      return { success: false, error: 'Error granting folder access: ' + error.message };
    }
  }

  async createOneTimeShare(recordUid, durationSeconds = 86400, editable = false) {
    try {
      const expireIn = this._formatDuration(durationSeconds || 604800);
      const editableFlag = editable ? ' --editable' : '';
      const command = 'one-time-share create' + editableFlag + ' ' + recordUid + ' -e ' + expireIn;
      
      const result = await this._executeCommandAsync(command, 30);
      
      if (!result || result.status !== 'success') {
        return { success: false, error: result?.message || 'Failed to create one-time share' };
      }
      
      let shareUrl = null;
      const message = result.message;
      
      if (typeof message === 'string') {
        if (message.startsWith('http')) {
          shareUrl = message;
        } else {
          const urlMatch = message.match(/https:\/\/[^\s]+/);
          if (urlMatch) {
            shareUrl = urlMatch[0];
          }
        }
      } else if (Array.isArray(message)) {
        for (const msg of message) {
          if (typeof msg === 'string' && msg.includes('https://')) {
            const urlMatch = msg.match(/https:\/\/[^\s]+/);
            if (urlMatch) {
              shareUrl = urlMatch[0];
              break;
            }
          }
        }
      }
      
      shareUrl = shareUrl || result.url || result.share_url || result.link;
      
      if (!shareUrl) {
        return {
          success: false,
          error: 'Share link created but URL not found in response',
          rawResponse: result,
        };
      }
      
      const expiresAt = new Date(Date.now() + (durationSeconds || 604800) * 1000);
      
      return {
        success: true,
        shareUrl: shareUrl,
        expiresAt: expiresAt.toISOString(),
        duration: expireIn,
      };
    } catch (error) {
      return { success: false, error: 'Error creating one-time share: ' + error.message };
    }
  }

  // ==================== PEDM Operations ====================

  async syncPedmData() {
    try {
      const result = await this._executeCommandAsync('pedm sync-down', 30);
      return result?.status === 'success';
    } catch (error) {
      console.error('[KeeperClient] PEDM sync failed:', error.message);
      return false;
    }
  }

  async getPendingPedmRequests() {
    try {
      await this.syncPedmData();
      
      const command = 'pedm approval list --type pending --format=json';
      const result = await this._executeCommandAsync(command, 30);
      
      if (!result || result.status !== 'success') {
        return [];
      }
      
      const data = result.data;
      if (!data || !Array.isArray(data)) {
        return [];
      }
      
      return data;
    } catch (error) {
      console.error('[KeeperClient] Error getting PEDM requests:', error.message);
      return [];
    }
  }

  async approvePedmRequest(approvalUid) {
    try {
      const command = 'pedm approval action --approve ' + approvalUid;
      const result = await this._executeCommandAsync(command, 15);
      
      if (result?.status === 'success') {
        return { success: true };
      } else {
        return { success: false, error: result?.message || 'Failed to approve PEDM request' };
      }
    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  async denyPedmRequest(approvalUid) {
    try {
      const command = 'pedm approval action --deny ' + approvalUid;
      const result = await this._executeCommandAsync(command, 15);
      
      if (result?.status === 'success') {
        return { success: true };
      } else {
        return { success: false, error: result?.message || 'Failed to deny PEDM request' };
      }
    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  // ==================== Device Approval Operations ====================

  async getPendingDeviceApprovals() {
    try {
      const command = 'device-approve --reload --format=json';
      const result = await this._executeCommandAsync(command, 30);
      
      if (!result || result.status !== 'success') {
        return [];
      }
      
      const data = result.data;
      if (!data || !Array.isArray(data)) {
        return [];
      }
      
      return data;
    } catch (error) {
      console.error('[KeeperClient] Error getting device approvals:', error.message);
      return [];
    }
  }

  async approveDevice(deviceId) {
    try {
      const command = 'device-approve --approve ' + deviceId;
      const result = await this._executeCommandAsync(command, 15);
      
      if (result?.status === 'success') {
        const message = result.message || '';
        if (typeof message === 'string' && message.toLowerCase().includes('no pending devices')) {
          return { success: false, alreadyHandled: true, error: 'Device request was already processed' };
        }
        return { success: true };
      } else {
        return { success: false, error: result?.message || 'Failed to approve device' };
      }
    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  async denyDevice(deviceId) {
    try {
      const command = 'device-approve --deny ' + deviceId;
      const result = await this._executeCommandAsync(command, 15);
      
      if (result?.status === 'success') {
        const message = result.message || '';
        if (typeof message === 'string' && message.toLowerCase().includes('no pending devices')) {
          return { success: false, alreadyHandled: true, error: 'Device request was already processed' };
        }
        return { success: true };
      } else {
        return { success: false, error: result?.message || 'Failed to deny device' };
      }
    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  // ==================== Helper Methods ====================

  async _executeCommandAsync(command, maxWait = 15) {
    try {
      console.log('[KeeperClient] Executing:', command);
      
      const response = await this.client.post('/executecommand-async', { command });
      
      if (response.status !== 202) {
        console.error('[KeeperClient] Command submission failed:', response.status);
        return null;
      }
      
      const requestId = response.data?.request_id;
      if (!requestId) {
        console.error('[KeeperClient] No request_id in response');
        return null;
      }
      
      return await this._pollForResult(requestId, maxWait);
    } catch (error) {
      console.error('[KeeperClient] Command execution error:', error.message);
      return null;
    }
  }

  async _pollForResult(requestId, maxWait = 15) {
    let pollInterval = 500;
    const maxPollInterval = 2000;
    let elapsed = 0;
    
    while (elapsed < maxWait * 1000) {
      try {
        const response = await this.client.get('/result/' + requestId);
        
        if (response.status === 200) {
          const result = response.data;
          const status = result?.status;
          
          if (status === 'success') {
            return result;
          } else if (status === 'error') {
            console.error('[KeeperClient] Command failed:', result.message);
            return result;
          }
        } else if (response.status === 400) {
          return { status: 'error', message: 'Command execution failed', httpStatus: 400 };
        }
        
        await this._sleep(pollInterval);
        elapsed += pollInterval;
        pollInterval = Math.min(pollInterval * 1.5, maxPollInterval);
        
      } catch (error) {
        if (error.response?.status === 202) {
          await this._sleep(pollInterval);
          elapsed += pollInterval;
          pollInterval = Math.min(pollInterval * 1.5, maxPollInterval);
        } else {
          console.error('[KeeperClient] Polling error:', error.message);
          return null;
        }
      }
    }
    
    console.warn('[KeeperClient] Polling timed out after ' + maxWait + ' seconds');
    return null;
  }

  _formatDuration(seconds) {
    if (!seconds || seconds <= 0) {
      return '7d';
    }
    
    if (seconds < 3600) {
      const minutes = Math.max(1, Math.floor(seconds / 60));
      return minutes + 'mi';
    } else if (seconds < 86400) {
      const hours = Math.max(1, Math.floor(seconds / 3600));
      return hours + 'h';
    } else if (seconds < 2592000) {
      const days = Math.max(1, Math.floor(seconds / 86400));
      return days + 'd';
    } else if (seconds < 31536000) {
      const months = Math.max(1, Math.floor(seconds / 2592000));
      return months + 'mo';
    } else {
      const years = Math.max(1, Math.floor(seconds / 31536000));
      return years + 'y';
    }
  }

  _parseSearchResults(data, type, limit) {
    if (!data || !Array.isArray(data)) {
      return [];
    }
    
    const results = [];
    
    for (const item of data) {
      if (results.length >= limit) break;
      
      if (type === 'record') {
        let recordType = 'login';
        let notes = '';
        
        if (item.details) {
          const parts = item.details.split(', ');
          for (const part of parts) {
            if (part.startsWith('Type: ')) {
              recordType = part.replace('Type: ', '').trim();
            } else if (part.startsWith('Description: ')) {
              notes = part.replace('Description: ', '').trim();
            }
          }
        }
        
        results.push({
          uid: item.uid,
          title: item.name,
          recordType: recordType,
          notes: notes,
        });
      } else {
        results.push({
          uid: item.uid,
          name: item.name,
          folderType: item.type || 'shared_folder',
        });
      }
    }
    
    return results;
  }

  _formatError(message) {
    if (!message) return 'Unknown error';
    
    if (Array.isArray(message)) {
      message = message.join('\n');
    }
    
    const lower = message.toLowerCase();
    
    if (lower.includes('time-limited access') && lower.includes('re-share')) {
      return 'User already has temporary access that conflicts with this permission. Remove existing access first.';
    }
    
    if (lower.includes('already') && (lower.includes('shared') || lower.includes('access'))) {
      return 'User already has existing permissions that conflict. Revoke existing access first.';
    }
    
    return message;
  }

  _sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

// Export singleton instance
let keeperClient;
try {
  keeperClient = new KeeperClient();
} catch (error) {
  console.warn('[KeeperClient] Could not initialize with config, creating with defaults');
  keeperClient = {
    healthCheck: async () => false,
    searchRecords: async () => [],
    searchFolders: async () => [],
    getRecordByUid: async () => null,
    getFolderByUid: async () => null,
    grantRecordAccess: async () => ({ success: false, error: 'Client not initialized' }),
    grantFolderAccess: async () => ({ success: false, error: 'Client not initialized' }),
    createOneTimeShare: async () => ({ success: false, error: 'Client not initialized' }),
    getPendingPedmRequests: async () => [],
    approvePedmRequest: async () => ({ success: false, error: 'Client not initialized' }),
    denyPedmRequest: async () => ({ success: false, error: 'Client not initialized' }),
    getPendingDeviceApprovals: async () => [],
    approveDevice: async () => ({ success: false, error: 'Client not initialized' }),
    denyDevice: async () => ({ success: false, error: 'Client not initialized' }),
  };
}

module.exports = keeperClient;
module.exports.KeeperClient = KeeperClient;
