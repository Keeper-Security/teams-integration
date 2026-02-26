/**
 * Keeper Commander Service Mode API Client
 */

const axios = require('axios');
const config = require('../config');
const { createLogger } = require('./logger');

const log = createLogger('KeeperClient');

/**
 * Sanitize search query to prevent command injection.
 * Removes dangerous shell characters that could be used for injection attacks.
 * Used for SEARCH queries only - not for record field values.
 * @param {string} query - User input query
 * @returns {string} - Sanitized query
 */
function sanitizeSearchQuery(query) {
  if (!query) return '';
  
  // Characters that could be used for command/shell injection
  // Matches Slack implementation for consistency
  const dangerousChars = /[;|&$`(){}[\]!\\<>"'\n\r\x00]/g;
  return query.replace(dangerousChars, '').trim();
}

/**
 * Shell escape a string for SEARCH queries.
 * First sanitizes (strips dangerous chars), then wraps in quotes if needed.
 * @param {string} str - Search query string
 * @returns {string} - Sanitized and escaped string
 */
function shellEscapeSearch(str) {
  if (!str) return '""';
  // First sanitize, then escape
  const sanitized = sanitizeSearchQuery(str);
  if (!sanitized) return '""';
  if (/[\s]/.test(sanitized)) {
    return '"' + sanitized + '"';
  }
  return sanitized;
}

/**
 * Shell escape a string for record field values (title, login, password, notes, url).
 * Properly escapes special characters WITHOUT stripping them.
 * This allows users to have passwords/titles with special characters.
 * @param {string} str - Field value string
 * @returns {string} - Properly escaped string
 */
function shellEscape(str) {
  if (!str) return '""';
  
  // Convert to string if not already
  const s = String(str);
  
  // Always wrap in single quotes and escape any single quotes within
  // This is the safest approach for shell escaping (matches Python's shlex.quote)
  // Single quotes preserve everything literally except single quotes themselves
  return "'" + s.replace(/'/g, "'\"'\"'") + "'";
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

  // updateCredentials(serviceUrl, apiKey = null) {
  //   this.baseUrl = serviceUrl.replace(/\/$/, '');
  //   this.apiKey = apiKey;
    
  //   this.client = axios.create({
  //     baseURL: this.baseUrl,
  //     timeout: 30000,
  //     headers: {
  //       'Content-Type': 'application/json',
  //       ...(apiKey && { 'api-key': apiKey }),
  //     },
  //   });
    
  //   console.log('[KeeperClient] Credentials updated:', this.baseUrl);
  // }

  // ==================== Health & Connection ====================

  async healthCheck() {
    try {
      const response = await this.client.get('/queue/status', { timeout: 5000 });
      return response.status === 200;
    } catch (error) {
      log.error('Health check failed', error.message);
      return false;
    }
  }

  // ==================== Search Operations ====================

  async searchRecords(query, limit = 20) {
    try {
      const command = 'search -c r ' + shellEscapeSearch(query) + ' --format=json';
      const result = await this._executeCommandAsync(command, 30);
      
      if (!result || result.status !== 'success') {
        log.debug('Search records failed', result?.message);
        return [];
      }
      
      return this._parseSearchResults(result.data, 'record', limit);
    } catch (error) {
      log.error('Error searching records', error.message);
      return [];
    }
  }

  async searchFolders(query, limit = 20) {
    try {
      const command = 'search -c s ' + shellEscapeSearch(query) + ' --format=json';
      const result = await this._executeCommandAsync(command, 30);
      
      if (!result || result.status !== 'success') {
        log.debug('Search folders failed', result?.message);
        return [];
      }
      
      return this._parseSearchResults(result.data, 'folder', limit);
    } catch (error) {
      log.error('Error searching folders', error.message);
      return [];
    }
  }

  async getRecordByUid(recordUid) {
    try {
      const command = 'search -c r ' + shellEscapeSearch(recordUid) + ' --format=json';
      const result = await this._executeCommandAsync(command, 30);
      
      if (!result || result.status !== 'success') {
        log.debug('Get record failed', result?.message);
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
      log.error('Error getting record', error.message);
      return null;
    }
  }

  async getFolderByUid(folderUid) {
    try {
      // Use -c s to filter for shared folders only (not records)
      const command = 'search -c s ' + shellEscapeSearch(folderUid) + ' --format=json';
      const result = await this._executeCommandAsync(command, 30);
      
      if (!result || result.status !== 'success') {
        log.debug('Get folder failed', result?.message);
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
      log.error('Error getting folder', error.message);
      return null;
    }
  }

  /**
   * Get the owner email of a record
   * @param {string} recordUid - Record UID
   * @returns {Promise<string|null>} - Owner email or null if not found
   */
  async getRecordOwner(recordUid) {
    try {
      const command = 'get --format=json ' + shellEscapeSearch(recordUid);
      const result = await this._executeCommandAsync(command, 30);
      
      if (!result || result.status !== 'success') {
        log.debug('Get record details failed', result?.message);
        return null;
      }
      
      const data = result.data;
      if (!data) {
        log.debug('No data in get result for record', recordUid);
        return null;
      }
      
      // Check user_permissions for owner
      const userPermissions = data.user_permissions || [];
      for (const userPerm of userPermissions) {
        if (userPerm.owner === true) {
          const ownerEmail = userPerm.username;
          log.debug('Found record owner:', ownerEmail);
          return ownerEmail;
        }
      }
      
      log.debug('No owner found in user_permissions for record', recordUid);
      return null;
    } catch (error) {
      log.error('Error getting record owner', error.message);
      return null;
    }
  }

  /**
   * Get the current permission level a user has on a record
   * @param {string} recordUid - Record UID
   * @param {string} userEmail - User's email
   * @returns {Promise<{permission: string, isOwner: boolean}|null>} - Permission info or null if no access
   */
  async getRecordUserPermission(recordUid, userEmail) {
    try {
      const command = 'get --format=json ' + shellEscapeSearch(recordUid);
      const result = await this._executeCommandAsync(command, 30);
      
      if (!result || result.status !== 'success') {
        log.debug('Get record details failed', result?.message);
        return null;
      }
      
      const data = result.data;
      if (!data) {
        log.debug('No data in get result for record', recordUid);
        return null;
      }
      
      // Check user_permissions for the specific user
      const userPermissions = data.user_permissions || [];
      const emailLower = userEmail.toLowerCase();
      
      for (const userPerm of userPermissions) {
        const permEmail = (userPerm.username || userPerm.email || '').toLowerCase();
        if (permEmail === emailLower) {
          const isOwner = userPerm.owner === true;
          const canEdit = userPerm.editable === true || userPerm.can_edit === true;
          const canShare = userPerm.shareable === true || userPerm.can_share === true;
          
          let permission;
          if (isOwner) {
            permission = 'owner';
          } else if (canEdit && canShare) {
            permission = 'edit_and_share';
          } else if (canShare) {
            permission = 'can_share';
          } else if (canEdit) {
            permission = 'can_edit';
          } else {
            permission = 'view_only';
          }
          
          log.debug('Found record permission for user', { userEmail, permission, isOwner });
          return { permission, isOwner };
        }
      }
      
      log.debug('No permission found for user on record', { recordUid, userEmail });
      return null;
    } catch (error) {
      log.error('Error getting record user permission', error.message);
      return null;
    }
  }

  /**
   * Get the current permission level a user has on a folder
   * @param {string} folderUid - Folder UID
   * @param {string} userEmail - User's email
   * @returns {Promise<{permission: string, manageUsers: boolean, manageRecords: boolean}|null>} - Permission info or null
   */
  async getFolderUserPermission(folderUid, userEmail) {
    try {
      const command = 'get --format=json ' + shellEscapeSearch(folderUid);
      const result = await this._executeCommandAsync(command, 30);
      
      if (!result || result.status !== 'success') {
        log.debug('Get folder details failed', result?.message);
        return null;
      }
      
      const data = result.data;
      if (!data) {
        log.debug('No data in get result for folder', folderUid);
        return null;
      }
      
      // Check users array for the specific user
      const userPermissions = data.users || [];
      const emailLower = userEmail.toLowerCase();
      
      for (const userPerm of userPermissions) {
        const permEmail = (userPerm.username || userPerm.email || '').toLowerCase();
        if (permEmail === emailLower) {
          const manageUsers = userPerm.manage_users === true;
          const manageRecords = userPerm.manage_records === true;
          
          // Determine permission level string
          let permission;
          if (manageUsers && manageRecords) {
            permission = 'manage_all';
          } else if (manageUsers) {
            permission = 'can_manage_users';
          } else if (manageRecords) {
            permission = 'can_manage_records';
          } else {
            permission = 'no_permissions';
          }
          
          log.debug('Found folder permission for user', { userEmail, permission, manageUsers, manageRecords });
          return { permission, manageUsers, manageRecords };
        }
      }
      
      log.debug('No permission found for user on folder', { folderUid, userEmail });
      return null;
    } catch (error) {
      log.error('Error getting folder user permission', error.message);
      return null;
    }
  }

  /**
   * Check if current record permission is equal to or higher than requested
   * Permission hierarchy: view_only < can_edit < can_share < edit_and_share < owner
   */
  _hasEqualOrHigherRecordPermission(currentPermission, requestedPermission) {
    const hierarchy = {
      'view_only': 1,
      'can_edit': 2,
      'can_share': 3,
      'edit_and_share': 4,
      'owner': 5,
    };
    
    const currentLevel = hierarchy[currentPermission] || 0;
    const requestedLevel = hierarchy[requestedPermission] || 0;
    
    // Only return true if user has EXACT SAME permission (not higher)
    // This allows upgrading or downgrading permissions via approval
    return currentLevel === requestedLevel;
  }

  /**
   * Check if current folder permission is equal to or higher than requested
   * Permission hierarchy: no_permissions < can_manage_records < can_manage_users < manage_all
   */
  _hasEqualOrHigherFolderPermission(currentPermission, requestedPermission) {
    const hierarchy = {
      'no_permissions': 1,
      'can_manage_records': 2,
      'can_manage_users': 3,
      'manage_all': 4,
    };
    
    const currentLevel = hierarchy[currentPermission] || 0;
    const requestedLevel = hierarchy[requestedPermission] || 0;
    
    // Only return true if user has EXACT SAME permission (not higher)
    // This allows upgrading or downgrading permissions via approval
    return currentLevel === requestedLevel;
  }

  /**
   * Get human-readable label for record permission
   */
  _getRecordPermissionLabel(permission) {
    const labels = {
      'view_only': 'View Only',
      'can_edit': 'Can Edit',
      'can_share': 'Can Share',
      'edit_and_share': 'Can Edit & Share',
      'owner': 'Owner',
    };
    return labels[permission] || permission;
  }

  /**
   * Get human-readable label for folder permission
   */
  _getFolderPermissionLabel(permission) {
    const labels = {
      'no_permissions': 'No Permissions (View Only)',
      'can_manage_records': 'Can Manage Records',
      'can_manage_users': 'Can Manage Users',
      'manage_all': 'Manage Users & Records',
    };
    return labels[permission] || permission;
  }

  // ==================== Access Grant Operations ====================

  async grantRecordAccess(recordUid, userEmail, permission, durationSeconds = 86400) {
    try {
      // Check if the user is the record owner (cannot modify owner's permissions)
      const recordOwner = await this.getRecordOwner(recordUid);
      if (recordOwner && userEmail.toLowerCase() === recordOwner.toLowerCase()) {
        log.info('Cannot grant access to record owner', { recordUid, userEmail });
        return {
          success: false,
          error: `Cannot modify permissions for record owner (${userEmail}). The user already owns this record and has full access to it.`,
          isOwnerError: true,
        };
      }
      
      // Check if user already has equal or higher permission (skip for change_owner)
      if (permission !== 'change_owner') {
        const currentPermission = await this.getRecordUserPermission(recordUid, userEmail);
        if (currentPermission) {
          if (currentPermission.isOwner) {
            log.info('User is already the owner of this record', { recordUid, userEmail });
            return {
              success: false,
              alreadyHasAccess: true,
              currentPermission: 'owner',
              currentPermissionLabel: this._getRecordPermissionLabel('owner'),
              error: `User (${userEmail}) is already the owner of this record with full access.`,
            };
          }
          
          if (this._hasEqualOrHigherRecordPermission(currentPermission.permission, permission)) {
            log.info('User already has equal or higher permission on record', { 
              recordUid, userEmail, 
              currentPermission: currentPermission.permission, 
              requestedPermission: permission 
            });
            return {
              success: false,
              alreadyHasAccess: true,
              currentPermission: currentPermission.permission,
              currentPermissionLabel: this._getRecordPermissionLabel(currentPermission.permission),
              error: `User (${userEmail}) already has "${this._getRecordPermissionLabel(currentPermission.permission)}" access to this record.`,
            };
          }
        }
      }
      
      if (permission === 'change_owner') {
        const command = 'share-record ' + recordUid + ' -e ' + userEmail + ' -a owner --force';
        const result = await this._executeCommandAsync(command, 30);
        
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
      
      // Retry logic for share command - helps with newly created records
      const maxShareRetries = 3;
      const shareRetryDelays = [0, 2000, 3000]; // No delay first, then 2s, 3s
      
      for (let shareAttempt = 0; shareAttempt < maxShareRetries; shareAttempt++) {
        if (shareAttempt > 0) {
          log.info('Retrying share command, attempt ' + (shareAttempt + 1), { recordUid, userEmail });
          await new Promise(resolve => setTimeout(resolve, shareRetryDelays[shareAttempt]));
        }
        
        const result = await this._executeCommandAsync(command, 30);
        
        if (result?.status === 'success') {
          // Check if invitation was sent (user doesn't have Keeper account yet)
          const message = result?.message || '';
          const invitationSent = this._isInvitationSent(message);
          
          if (shareAttempt > 0) {
            log.info('Share succeeded on retry attempt ' + (shareAttempt + 1), { recordUid });
          }
          
          return {
            success: true,
            expiresAt: expiresAtStr,
            permission: permission,
            duration: isPermanent ? 'permanent' : 'temporary',
            invitationSent: invitationSent,
          };
        }
        
        // Check if the "error" is actually an invitation sent scenario
        // Commander sometimes returns invitation as an error
        const errorMsg = this._formatError(result?.message) || '';
        const errorField = result?.error || '';
        const combinedMessage = errorMsg + ' ' + errorField;
        
        if (this._isInvitationSent(combinedMessage)) {
          log.info('Record share invitation sent (detected in error response)', { recordUid, userEmail });
          return {
            success: true,
            expiresAt: 'Pending Invitation',
            permission: permission,
            duration: 'permanent',
            invitationSent: true,
          };
        }
        
        // Check if this is a retryable error
        const isRetryable = errorMsg.toLowerCase().includes('not found') || 
                           errorMsg.toLowerCase().includes('does not exist') ||
                           errorMsg.toLowerCase().includes('unknown') ||
                           errorMsg.toLowerCase().includes('error');
        
        if (!isRetryable || shareAttempt === maxShareRetries - 1) {
          // Non-retryable error or last attempt failed
          return { success: false, error: errorMsg || 'Failed to share record' };
        }
        
        log.debug('Share failed with retryable error, will retry', { error: errorMsg, attempt: shareAttempt + 1 });
      }
      
      return { success: false, error: 'Failed to share record after multiple attempts' };
    } catch (error) {
      return { success: false, error: 'Error granting record access: ' + error.message };
    }
  }

  async grantFolderAccess(folderUid, userEmail, permission, durationSeconds = 86400) {
    try {
      // Check if user already has permission on this folder
      const currentPermission = await this.getFolderUserPermission(folderUid, userEmail);
      if (currentPermission) {
        // Check if user has manage_all - cannot be downgraded
        if (currentPermission.manageUsers && currentPermission.manageRecords) {
          if (permission !== 'manage_all') {
            log.info('Cannot downgrade folder permission from manage_all', { folderUid, userEmail, requestedPermission: permission });
            return {
              success: false,
              error: `Cannot modify permissions for user (${userEmail}). The user already has "Manage Users and Records" permission on this folder and cannot be downgraded.`,
              isFullAccessError: true,
            };
          }
          // User already has manage_all and requesting manage_all - already has access
          log.info('User already has manage_all permission on folder', { folderUid, userEmail });
          return {
            success: false,
            alreadyHasAccess: true,
            currentPermission: currentPermission.permission,
            currentPermissionLabel: this._getFolderPermissionLabel(currentPermission.permission),
            error: `User (${userEmail}) already has "${this._getFolderPermissionLabel(currentPermission.permission)}" access to this folder.`,
          };
        }
        
        // Check if user already has equal or higher permission
        if (this._hasEqualOrHigherFolderPermission(currentPermission.permission, permission)) {
          log.info('User already has equal or higher permission on folder', { 
            folderUid, userEmail, 
            currentPermission: currentPermission.permission, 
            requestedPermission: permission 
          });
          return {
            success: false,
            alreadyHasAccess: true,
            currentPermission: currentPermission.permission,
            currentPermissionLabel: this._getFolderPermissionLabel(currentPermission.permission),
            error: `User (${userEmail}) already has "${this._getFolderPermissionLabel(currentPermission.permission)}" access to this folder.`,
          };
        }
      }
      
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
      
      const result = await this._executeCommandAsync(command, 30);
      
      if (result?.status === 'success') {
        // Check if invitation was sent (user doesn't have Keeper account yet)
        const message = result?.message || '';
        const invitationSent = this._isInvitationSent(message);
        
        return {
          success: true,
          expiresAt: expiresAtStr,
          permission: permission,
          duration: isPermanent ? 'permanent' : 'temporary',
          invitationSent: invitationSent,
        };
      } else {
        // Check if the "error" is actually an invitation sent scenario
        // Commander sometimes returns invitation as an error (http_status 400)
        const errorMsg = this._formatError(result?.message);
        const errorField = result?.error || '';
        const combinedMessage = errorMsg + ' ' + errorField;
        
        if (this._isInvitationSent(combinedMessage)) {
          log.info('Folder share invitation sent (detected in error response)', { folderUid, userEmail });
          return {
            success: true,
            expiresAt: 'Pending Invitation',
            permission: permission,
            duration: 'permanent',
            invitationSent: true,
          };
        }
        
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

  // ==================== Record Creation ====================

  /**
   * Create a new record in Keeper vault
   * @param {Object} params - Record parameters
   * @param {string} params.title - Record title (required)
   * @param {string} params.login - Login/username
   * @param {string} params.password - Password (use '$GEN' to auto-generate)
   * @param {string} params.url - Website URL
   * @param {string} params.notes - Notes
   * @param {boolean} params.generatePassword - Whether to auto-generate password
   * @param {string} params.selfDestructDuration - Self-destruct duration (e.g., '1h', '24h', '7d', '30d', '90d')
   * @returns {Promise<Object>} - Result with success, recordUid, etc.
   */
  async createRecord({ title, login, password, url, notes, generatePassword = false, selfDestructDuration = null }) {
    try {
      if (!title || !title.trim()) {
        return { success: false, error: 'Title is required' };
      }

      // Build the record-add command
      const commandParts = ['record-add'];
      
      // Add record type
      commandParts.push('--record-type login');
      
      // Add title
      commandParts.push('--title ' + shellEscape(title));
      
      // Add notes if provided
      if (notes && notes.trim()) {
        const notesForCli = notes.replace(/\n/g, '\\n');
        commandParts.push('--notes ' + shellEscape(notesForCli));
      }
      
      // Add self-destruct duration if provided
      if (selfDestructDuration && selfDestructDuration.trim()) {
        commandParts.push('--self-destruct ' + selfDestructDuration);
      }
      
      // Add login if provided
      if (login && login.trim()) {
        commandParts.push('login=' + shellEscape(login));
      }
      
      // Add password
      if (password && password.trim()) {
        if (password === '$GEN') {
          commandParts.push('password=$GEN');
        } else {
          commandParts.push('password=' + shellEscape(password));
        }
      } else if (generatePassword) {
        commandParts.push('password=$GEN');
      }
      
      // Add URL if provided
      if (url && url.trim()) {
        commandParts.push('url=' + shellEscape(url));
      }
      
      const command = commandParts.join(' ');
      log.debug('Creating record with command', command.replace(/password=[^\s]+/, 'password=***'));
      
      const result = await this._executeCommandAsync(command, 20);
      
      if (!result || result.status !== 'success') {
        const errorMsg = this._formatError(result?.message);
        return { success: false, error: 'Failed to create record: ' + errorMsg };
      }
      
      log.debug('Record created successfully, searching for UID...');
      
      // record-add doesn't return the UID, so we need to search for it
      // Wait for the record to be indexed - use longer wait with retry
      // This addresses timing issues where the record isn't immediately available
      
      const maxRetries = 3;
      const waitTimes = [2000, 3000, 4000]; // Increasing wait times: 2s, 3s, 4s
      
      for (let attempt = 0; attempt < maxRetries; attempt++) {
        await new Promise(resolve => setTimeout(resolve, waitTimes[attempt]));
        
        try {
          const searchCommand = 'search ' + shellEscapeSearch(title) + ' --format=json';
          const searchResult = await this._executeCommandAsync(searchCommand, 10);
          
          if (searchResult && searchResult.status === 'success' && searchResult.data) {
            const data = Array.isArray(searchResult.data) ? searchResult.data : [];
            
            if (data.length > 0) {
              // Find the record with matching title (most recently created should be first)
              const matchingRecord = data.find(r => r.title === title) || data[0];
              const recordUid = matchingRecord.uid || matchingRecord.record_uid;
              
              if (recordUid) {
                log.info('Found created record UID after attempt ' + (attempt + 1), { recordUid, title });
                return {
                  success: true,
                  recordUid: recordUid,
                  title: title,
                  generatedPassword: generatePassword || password === '$GEN',
                  isNewlyCreated: true, // Flag to indicate this is a newly created record
                };
              }
            }
          }
          
          log.debug('Record not found yet, attempt ' + (attempt + 1) + ' of ' + maxRetries);
        } catch (searchError) {
          log.debug('Search error on attempt ' + (attempt + 1), searchError.message);
        }
      }
      
      log.warn('Record created but UID not found after ' + maxRetries + ' attempts');
      return {
        success: false,
        error: 'Record created but UID could not be retrieved. The record exists in your vault but the approval flow cannot continue automatically. Please try searching for the record manually.',
      };
    } catch (error) {
      log.error('Exception in createRecord', error);
      return { success: false, error: 'Error creating record: ' + error.message };
    }
  }

  // ==================== PEDM Operations ====================

  async syncPedmData() {
    try {
      const result = await this._executeCommandAsync('epm sync-down', 30);
      return result?.status === 'success';
    } catch (error) {
      log.error('PEDM sync failed', error.message);
      return false;
    }
  }

  async getPendingPedmRequests() {
    try {
      await this.syncPedmData();
      
      const command = 'epm approval list --type pending --format=json';
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
      log.error('Error getting PEDM requests', error.message);
      return [];
    }
  }

  async approvePedmRequest(approvalUid) {
    try {
      const command = 'epm approval action --approve ' + approvalUid;
      const result = await this._executeCommandAsync(command, 30);
      
      if (result?.status === 'success') {
        // Check message for "already processed" indicators even on success status
        const message = result.message || '';
        const messageLower = typeof message === 'string' ? message.toLowerCase() : '';
        if (messageLower.includes('does not exist') || 
            messageLower.includes('cannot be modified') ||
            messageLower.includes('not found') ||
            messageLower.includes('no output')) {
          log.info('EPM request already processed detected in success response', { approvalUid });
          return { 
            success: false, 
            error: 'This request has already been approved or denied.',
            already_processed: true 
          };
        }
        return { success: true };
      } else {
        const errorMsg = this._formatError(result?.message) || 'Failed to approve PEDM request';
        
        // Check if this is an "already processed" error from message
        if (errorMsg.includes('does not exist or cannot be modified') || 
            errorMsg.includes('Approval request does not exist') ||
            errorMsg.includes('not found') ||
            errorMsg.includes('already')) {
          return { 
            success: false, 
            error: 'This request has already been approved or denied.',
            already_processed: true 
          };
        }
        
        // Optimized fallback: if result is null (500/timeout), check if request is still pending
        if (!result) {
          log.info('EPM approve failed with null result, checking if still pending', { approvalUid });
          const pendingRequests = await this.getPendingPedmRequests();
          const stillPending = pendingRequests.some(r => 
            r.approval_uid === approvalUid || r.approvalUid === approvalUid || r.uid === approvalUid
          );
          
          if (!stillPending) {
            log.info('EPM request not in pending list - already processed elsewhere', { approvalUid });
            return { 
              success: false, 
              error: 'This request has already been approved or denied.',
              already_processed: true 
            };
          }
        }
        
        return { success: false, error: errorMsg };
      }
    } catch (error) {
      return { success: false, error: error.message };
    }
  }

  async denyPedmRequest(approvalUid) {
    try {
      const command = 'epm approval action --deny ' + approvalUid;
      const result = await this._executeCommandAsync(command, 30);
      
      if (result?.status === 'success') {
        // Check message for "already processed" indicators even on success status
        const message = result.message || '';
        const messageLower = typeof message === 'string' ? message.toLowerCase() : '';
        if (messageLower.includes('does not exist') || 
            messageLower.includes('cannot be modified') ||
            messageLower.includes('not found') ||
            messageLower.includes('no output')) {
          log.info('EPM request already processed detected in success response', { approvalUid });
          return { 
            success: false, 
            error: 'This request has already been approved or denied.',
            already_processed: true 
          };
        }
        return { success: true };
      } else {
        const errorMsg = this._formatError(result?.message) || 'Failed to deny PEDM request';
        
        // Check if this is an "already processed" error from message
        if (errorMsg.includes('does not exist or cannot be modified') || 
            errorMsg.includes('Approval request does not exist') ||
            errorMsg.includes('not found') ||
            errorMsg.includes('already')) {
          return { 
            success: false, 
            error: 'This request has already been approved or denied.',
            already_processed: true 
          };
        }
        
        // Optimized fallback: if result is null (500/timeout), check if request is still pending
        if (!result) {
          log.info('EPM deny failed with null result, checking if still pending', { approvalUid });
          const pendingRequests = await this.getPendingPedmRequests();
          const stillPending = pendingRequests.some(r => 
            r.approval_uid === approvalUid || r.approvalUid === approvalUid || r.uid === approvalUid
          );
          
          if (!stillPending) {
            log.info('EPM request not in pending list - already processed elsewhere', { approvalUid });
            return { 
              success: false, 
              error: 'This request has already been approved or denied.',
              already_processed: true 
            };
          }
        }
        
        return { success: false, error: errorMsg };
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
      log.error('Error getting device approvals', error.message);
      return [];
    }
  }

  async approveDevice(deviceId) {
    try {
      // First check if device is still pending (prevents approving already-processed requests)
      const pendingDevices = await this.getPendingDeviceApprovals();
      const isPending = pendingDevices.some(d => 
        d.device_id === deviceId || d.deviceId === deviceId || d.id === deviceId
      );
      
      if (!isPending) {
        log.info('Device not in pending list - already processed elsewhere', { deviceId });
        return { 
          success: false, 
          already_processed: true, 
          error: 'This request has already been approved or denied from another platform.' 
        };
      }
      
      // Device is pending, proceed with approval
      const command = 'device-approve --approve ' + deviceId;
      const result = await this._executeCommandAsync(command, 30);
      
      log.info('approveDevice result:', JSON.stringify(result));
      
      if (result?.status === 'success') {
        // Also check message and data for "no pending" indicators
        const message = result.message || '';
        const dataStr = typeof result.data === 'string' ? result.data : JSON.stringify(result.data || '');
        const combinedText = (message + ' ' + dataStr).toLowerCase();
        
        // Check if this is an "already processed" response
        if (combinedText.includes('no pending') || 
            combinedText.includes('there are no pending') ||
            combinedText.includes('already processed') ||
            combinedText.includes('already approved') ||
            combinedText.includes('already denied') ||
            combinedText.includes('no output')) {
          log.info('Device already processed detected in success response', deviceId);
          return { success: false, already_processed: true, error: 'This request has already been approved or denied from another platform.' };
        }
        return { success: true };
      } else {
        const errorMsg = this._formatError(result?.message) || 'Failed to approve device';
        const errorLower = errorMsg.toLowerCase();
        
        // Check if this is an "already processed" error (case-insensitive)
        if (errorLower.includes('no pending') || 
            errorLower.includes('not found') ||
            errorLower.includes('already') ||
            errorLower.includes('does not exist') ||
            errorLower.includes('invalid device')) {
          log.info('Device already processed detected in error response', deviceId);
          return { 
            success: false, 
            error: 'This request has already been approved or denied from another platform.',
            already_processed: true 
          };
        }
        
        return { success: false, error: errorMsg };
      }
    } catch (error) {
      log.error('approveDevice exception:', error.message);
      return { success: false, error: error.message };
    }
  }

  async denyDevice(deviceId) {
    try {
      // Optimized: try operation first, only check pending list on failure
      const command = 'device-approve --deny ' + deviceId;
      const result = await this._executeCommandAsync(command, 30);
      
      log.info('denyDevice result:', JSON.stringify(result));
      
      if (result?.status === 'success') {
        const message = result.message || '';
        const messageLower = typeof message === 'string' ? message.toLowerCase() : '';
        
        // Check if this is an "already processed" response
        if (messageLower.includes('no pending') || 
            messageLower.includes('there are no pending') ||
            messageLower.includes('already processed') ||
            messageLower.includes('already approved') ||
            messageLower.includes('already denied') ||
            messageLower.includes('no output')) {
          log.info('Device already processed detected in success response', deviceId);
          return { success: false, already_processed: true, error: 'This request has already been approved or denied from another platform.' };
        }
        return { success: true };
      } else {
        const errorMsg = this._formatError(result?.message) || 'Failed to deny device';
        const errorLower = errorMsg.toLowerCase();
        
        // Check if this is an "already processed" error (case-insensitive)
        if (errorLower.includes('no pending') || 
            errorLower.includes('not found') ||
            errorLower.includes('already') ||
            errorLower.includes('does not exist') ||
            errorLower.includes('invalid device')) {
          log.info('Device already processed detected in error response', deviceId);
          return { 
            success: false, 
            error: 'This request has already been approved or denied from another platform.',
            already_processed: true 
          };
        }
        
        // Optimized fallback: if result is null (500/timeout), check pending list
        if (!result) {
          log.info('Device deny failed with null result, checking if still pending', { deviceId });
          const pendingDevices = await this.getPendingDeviceApprovals();
          const stillPending = pendingDevices.some(d => 
            d.device_id === deviceId || d.deviceId === deviceId || d.id === deviceId
          );
          
          if (!stillPending) {
            log.info('Device not in pending list - already processed elsewhere', { deviceId });
            return { 
              success: false, 
              already_processed: true, 
              error: 'This request has already been approved or denied from another platform.' 
            };
          }
        }
        
        return { success: false, error: errorMsg };
      }
    } catch (error) {
      log.error('denyDevice exception:', error.message);
      return { success: false, error: error.message };
    }
  }

  // ==================== Helper Methods ====================

  async _executeCommandAsync(command, maxWait = 30) {
    try {
      log.debug('Executing command', command);
      
      const response = await this.client.post('/executecommand-async', { command });
      
      if (response.status !== 202) {
        log.error('Command submission failed', response.status);
        return null;
      }
      
      const requestId = response.data?.request_id;
      if (!requestId) {
        log.error('No request_id in response');
        return null;
      }
      
      return await this._pollForResult(requestId, maxWait);
    } catch (error) {
      log.error('Command execution error', error.message);
      return null;
    }
  }

  async _pollForResult(requestId, maxWait = 30) {
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
            log.error('Command failed', result.message);
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
        } else if (error.response?.status === 400 || error.response?.status === 500) {
          // Capture error response data - may contain invitation messages
          const errorData = error.response?.data || {};
          log.debug('Polling error with response data', { status: error.response?.status, data: errorData });
          return {
            status: 'error',
            message: errorData.message || errorData.error || error.message,
            error: errorData.error || errorData.message || error.message,
            httpStatus: error.response?.status,
            data: errorData
          };
        } else {
          log.error('Polling error', error.message);
          return null;
        }
      }
    }
    
    log.warn('Polling timed out after ' + maxWait + ' seconds');
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

  /**
   * Check if the share command result indicates an invitation was sent
   * (user doesn't have a Keeper account yet)
   */
  _isInvitationSent(message) {
    if (!message) return false;
    
    if (Array.isArray(message)) {
      message = message.join(' ');
    }
    
    const lower = String(message).toLowerCase();
    
    // Primary indicators from Keeper Commander (matches Slack implementation)
    if (lower.includes('invitation has been sent') || 
        lower.includes('repeat this command when invitation is accepted')) {
      return true;
    }
    
    // Additional indicators that an invitation was sent instead of immediate share
    return lower.includes('invitation') || 
           lower.includes('invite') ||
           lower.includes('pending share') ||
           lower.includes('not a keeper user') ||
           lower.includes('does not have a keeper') ||
           lower.includes('user not found') ||
           lower.includes('email will be sent') ||
           lower.includes('share pending');
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
  log.warn('Could not initialize with config, creating with defaults');
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
