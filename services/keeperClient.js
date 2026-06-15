/**
 * Keeper Commander Service Mode API Client
 */

const axios = require('axios');
const { getConfig } = require('../config');
const { createLogger } = require('./logger');
const { isNsfType, isValidNsfRole, isNsfRecordCategory } = require('../utils/helpers');

const log = createLogger('KeeperClient');

/**
 * Sanitize search query to prevent command injection.
 * @param {string} query - User input query
 * @returns {string} - Sanitized query
 */
function sanitizeSearchQuery(query) {
  if (!query) return '';
  const dangerousChars = /[;|&$`(){}[\]!\\<>"'\n\r\x00]/g;
  return query.replace(dangerousChars, '').trim();
}

/**
 * Shell escape a string for SEARCH queries.
 * @param {string} str - Search query string
 * @returns {string} - Sanitized and escaped string
 */
function shellEscapeSearch(str) {
  if (!str) return '""';
  const sanitized = sanitizeSearchQuery(str);
  if (!sanitized) return '""';
  if (/[\s]/.test(sanitized)) {
    return '"' + sanitized + '"';
  }
  return sanitized;
}

/**
 * Shell escape a string for record field values (title, login, password, notes, url).
 * @param {string} str - Field value string
 * @returns {string} - Properly escaped string
 */
function shellEscape(str) {
  if (!str) return '""';
  const s = String(str);
  return "'" + s.replace(/'/g, "'\"'\"'") + "'";
}

class KeeperClient {
  constructor() {
    // Get current config (with KSM data if initialized)
    const config = getConfig();
    this.baseUrl = (config.keeper?.serviceUrl || 'http://localhost:8900/api/v2/').replace(/\/$/, '');
    this.apiKey = config.keeper?.apiKey;
    
    this.client = axios.create({
      baseURL: this.baseUrl,
      timeout: 30000,
      headers: {
        'Content-Type': 'application/json',
        ...(this.apiKey && { 'api-key': this.apiKey }),
      },
    });
    
    log.info('KeeperClient initialized', { baseUrl: this.baseUrl });
  }



  async healthCheck() {
    try {
      const response = await this.client.get('/queue/status', { timeout: 5000 });
      return response.status === 200;
    } catch (error) {
      log.error('Health check failed', error.message);
      return false;
    }
  }


  async searchRecords(query, limit = 20) {
    try {
      const command = 'search -c r ' + shellEscape(query) + ' --format=json';
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
      const command = 'search -c s,d ' + shellEscape(query) + ' --format=json';
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
      let recordCategory = '';
      
      if (item.details) {
        const parts = item.details.split(', ');
        for (const part of parts) {
          if (part.startsWith('Type: ')) {
            recordType = part.replace('Type: ', '').trim();
          } else if (part.startsWith('Description: ')) {
            notes = part.replace('Description: ', '').trim();
          } else if (part.startsWith('Record Category: ')) {
            recordCategory = part.replace('Record Category: ', '').trim();
          }
        }
      }
      
      return {
        uid: item.uid || recordUid,
        title: item.name || 'Untitled Record',
        recordType: itemType.includes('folder') ? itemType : recordType,
        folderUid: null,
        notes: notes,
        // NSF records keep their normal type; "Record Category: Nested" is the
        // reliable signal. Keep the type check as a fallback.
        isNsf: isNsfRecordCategory(recordCategory) || isNsfType(itemType),
      };
    } catch (error) {
      log.error('Error getting record', error.message);
      return null;
    }
  }

  async getFolderByUid(folderUid) {
    try {
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
        isNsf: isNsfType(item.type),
      };
    } catch (error) {
      log.error('Error getting folder', error.message);
      return null;
    }
  }

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
    return currentLevel === requestedLevel;
  }

  _normalizeFolderPermission(permission) {
    const mapping = {
      'manage_records': 'can_manage_records',
      'manage_users': 'can_manage_users',
      'can_manage_records': 'can_manage_records',
      'can_manage_users': 'can_manage_users',
      'manage_all': 'manage_all',
      'no_permissions': 'no_permissions',
    };
    return mapping[permission] || permission;
  }


  _hasEqualOrHigherFolderPermission(currentPermission, requestedPermission) {
    const hierarchy = {
      'no_permissions': 1,
      'can_manage_records': 2,
      'can_manage_users': 3,
      'manage_all': 4,
    };
    
    // Normalize both permissions to ensure consistent comparison
    const normalizedCurrent = this._normalizeFolderPermission(currentPermission);
    const normalizedRequested = this._normalizeFolderPermission(requestedPermission);
    
    const currentLevel = hierarchy[normalizedCurrent] || 0;
    const requestedLevel = hierarchy[normalizedRequested] || 0;
    
    // Only return true if user has EXACT SAME permission (not higher)
    // This allows upgrading or downgrading permissions via approval
    return currentLevel === requestedLevel;
  }

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


  _getFolderPermissionLabel(permission) {
    const labels = {
      'no_permissions': 'No Permissions (View Only)',
      'can_manage_records': 'Can Manage Records',
      'can_manage_users': 'Can Manage Users',
      'manage_records': 'Can Manage Records',
      'manage_users': 'Can Manage Users',
      'manage_all': 'Manage Users & Records',
    };
    return labels[permission] || permission;
  }

  async isPamUserFolder(folderUid) {
    if (!folderUid) return false;
    try {
      const command = 'list-sf ' + shellEscapeSearch(folderUid) + ' --roe-eligible --format=json';
      const result = await this._executeCommandAsync(command, 10);
      if (!result || result.status !== 'success') {
        log.debug('list-sf --roe-eligible non-success for folder', folderUid);
        return false;
      }
      const data = result.data;
      if (Array.isArray(data) && data.length > 0) {
        log.debug('Folder is rotate-on-expire eligible', { folderUid, entries: data.length });
        return true;
      }
      return false;
    } catch (error) {
      log.error('Exception in isPamUserFolder', { folderUid, error: error.message });
      return false;
    }
  }

  async grantRecordAccess(recordUid, userEmail, permission, durationSeconds = 86400, rotateOnExpire = false) {
    try {
      try { await this._executeCommandAsync('sync-down', 15); } catch (e) { log.warn('sync-down before share-record failed, proceeding', e.message); }

      // Check if the user is the record owner (cannot modify owner's permissions)
      // This is the only pre-check needed - aligns with Slack app implementation
      const recordOwner = await this.getRecordOwner(recordUid);
      if (recordOwner && userEmail.toLowerCase() === recordOwner.toLowerCase()) {
        log.info('Cannot grant access to record owner', { recordUid, userEmail });
        return {
          success: false,
          error: `Cannot modify permissions for record owner (${userEmail}). The user already owns this record and has full access to it.`,
          isOwnerError: true,
        };
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
        if (rotateOnExpire) {
          command += ' --rotate-on-expiration';
        }
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
          
          const pamRotateScheduled = !!(rotateOnExpire && !isPermanent && durationSeconds);
          return {
            success: true,
            expiresAt: expiresAtStr,
            permission: permission,
            duration: isPermanent ? 'permanent' : 'temporary',
            invitationSent: invitationSent,
            rotateOnExpire: pamRotateScheduled,
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
        
        // Detect rotation-not-configured error (non-retryable)
        const errLower = errorMsg.toLowerCase();
        if (
          errLower.includes('rotation must be already set') ||
          (errLower.includes('rotate') && errLower.includes('expiration') && errLower.includes('set on the record')) ||
          (errLower.includes('--rotate-on-expiration') && (errLower.includes('requires') || errLower.includes('ineligible')))
        ) {
          return {
            success: false,
            errorCode: 'pam_rotation_not_configured',
            error: 'Rotation is not configured on this PAM User record.\n\nSet up rotation (Gateway + rotation settings) on the record in the Keeper Vault first, or uncheck "Rotate credentials when access expires" and approve again.',
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

  async grantFolderAccess(folderUid, userEmail, permission, durationSeconds = 86400, rotateOnExpire = false) {
    try {
      try { await this._executeCommandAsync('sync-down', 15); } catch (e) { log.warn('sync-down before share-folder failed, proceeding', e.message); }

      // No permission pre-check needed - aligns with Slack app implementation
      // The grant command with revoke+retry pattern handles permission conflicts gracefully
      
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
        if (rotateOnExpire) {
          command += ' --rotate-on-expiration';
        }
      }
      
      command += ' -f';
      
      log.info('Executing folder grant command', { command, permission, isPermanent });
      const result = await this._executeCommandAsync(command, 30);
      log.info('Folder grant result', { status: result?.status, message: result?.message, error: result?.error });
      
      if (result?.status === 'success') {
        // Check if invitation was sent (user doesn't have Keeper account yet)
        const message = result?.message || '';
        const invitationSent = this._isInvitationSent(message);
        const pamRotateScheduled = !!(rotateOnExpire && !isPermanent && durationSeconds);
        
        return {
          success: true,
          expiresAt: expiresAtStr,
          permission: permission,
          duration: isPermanent ? 'permanent' : 'temporary',
          invitationSent: invitationSent,
          rotateOnExpire: pamRotateScheduled,
        };
      } else {
        // Check if the "error" is actually an invitation sent scenario
        // Commander sometimes returns invitation as an error (http_status 400)
        const errorMsg = this._formatError(result?.message);
        const errorField = result?.error || '';
        const combinedMessage = (errorMsg + ' ' + errorField).toLowerCase();
        
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
        
        // Detect rotation-not-configured error (non-retryable)
        if (
          combinedMessage.includes('rotation must be already set') ||
          (combinedMessage.includes('rotate') && combinedMessage.includes('expiration') && combinedMessage.includes('set on the record'))
        ) {
          return {
            success: false,
            errorCode: 'pam_rotation_not_configured',
            error: 'Rotation is not configured for this PAM User folder.\n\nSet up rotation (Gateway + rotation settings) on the records in this folder in the Keeper Vault first, or uncheck "Rotate credentials when access expires" and approve again.',
          };
        }

        // Check if this is a "User share failed" error - indicates existing access conflict
        // Try revoking existing access and retrying the grant
        const isUserShareFailed = combinedMessage.includes('user share') && combinedMessage.includes('failed');
        const isTimeConflict = combinedMessage.includes('time-limited') || combinedMessage.includes('already');
        
        if (isUserShareFailed || isTimeConflict) {
          log.info('Grant failed due to existing access, attempting revoke+grant', { folderUid, userEmail, error: errorMsg });
          
          try {
            // Revoke existing permission
            const revokeCmd = 'share-folder ' + folderUid + ' -e ' + userEmail + ' -a remove -f';
            const revokeResult = await this._executeCommandAsync(revokeCmd, 15);
            log.info('Revoke result', { status: revokeResult?.status, message: revokeResult?.message });
            
            // Wait for revoke to propagate
            await new Promise(resolve => setTimeout(resolve, 1500));
            
            // Retry the grant
            log.info('Retrying folder grant after revoke', { command });
            const retryResult = await this._executeCommandAsync(command, 30);
            log.info('Retry grant result', { status: retryResult?.status, message: retryResult?.message });
            
            if (retryResult?.status === 'success') {
              const message = retryResult?.message || '';
              const invitationSent = this._isInvitationSent(message);
              return {
                success: true,
                expiresAt: expiresAtStr,
                permission: permission,
                duration: isPermanent ? 'permanent' : 'temporary',
                invitationSent: invitationSent,
              };
            }
            
            // Check if retry resulted in invitation
            const retryMsg = (this._formatError(retryResult?.message) + ' ' + (retryResult?.error || '')).toLowerCase();
            if (this._isInvitationSent(retryMsg)) {
              return {
                success: true,
                expiresAt: 'Pending Invitation',
                permission: permission,
                duration: 'permanent',
                invitationSent: true,
              };
            }
            
            // Retry also failed
            return { success: false, error: this._formatError(retryResult?.message) || errorMsg };
          } catch (retryError) {
            log.error('Revoke+retry failed', { error: retryError.message });
            return { success: false, error: errorMsg };
          }
        }
        
        return { success: false, error: errorMsg };
      }
    } catch (error) {
      return { success: false, error: 'Error granting folder access: ' + error.message };
    }
  }

  /**
   * Grant a user access to a Nested Share Folder (NSF) record via `nsf-share-record`.
   * NSF uses a single role (viewer ... full-manager) instead of classic permission flags.
   * @param {string} recordUid - NSF record UID
   * @param {string} userEmail - Recipient email
   * @param {string} role - NSF role (viewer, share-manager, content-manager, content-share-manager, full-manager)
   * @param {number|null} durationSeconds - Access duration in seconds; null = permanent
   * @returns {Promise<object>} Result shaped like grantRecordAccess
   */
  async grantNsfRecordAccess(recordUid, userEmail, role, durationSeconds = 86400) {
    try {
      // Transfer Ownership is a distinct NSF action (records only) and is always
      // permanent. It uses `-a owner` instead of `-a grant -r <role>`.
      if (role === 'owner') {
        try { await this._executeCommandAsync('sync-down', 15); } catch (e) { log.warn('sync-down before nsf-share-record (owner) failed, proceeding', e.message); }

        const ownerCommand = 'nsf-share-record ' + recordUid + ' -e ' + userEmail + ' -a owner --force';
        log.info('Executing NSF record ownership transfer command', { command: ownerCommand });
        const ownerResult = await this._executeCommandAsync(ownerCommand, 30);
        log.info('NSF record ownership transfer result', { status: ownerResult?.status, message: ownerResult?.message, error: ownerResult?.error });

        if (ownerResult?.status === 'success') {
          return {
            success: true,
            expiresAt: 'N/A (Ownership Transfer)',
            permission: 'owner',
            nsfRole: 'owner',
            isNsf: true,
            duration: 'permanent',
            invitationSent: this._isInvitationSent(ownerResult?.message || ''),
          };
        }

        const ownerErr = this._formatError(ownerResult?.message) || '';
        const ownerErrField = ownerResult?.error || '';
        if (this._isInvitationSent((ownerErr + ' ' + ownerErrField).trim())) {
          return {
            success: true,
            expiresAt: 'Pending Invitation',
            permission: 'owner',
            nsfRole: 'owner',
            isNsf: true,
            duration: 'permanent',
            invitationSent: true,
          };
        }
        return { success: false, error: ownerErr || ownerErrField || 'Failed to transfer Nested Share record ownership' };
      }

      if (!isValidNsfRole(role)) {
        return { success: false, error: `Invalid Nested Share role: ${role}` };
      }

      try { await this._executeCommandAsync('sync-down', 15); } catch (e) { log.warn('sync-down before nsf-share-record failed, proceeding', e.message); }

      // Revoke any existing share for this user first so a re-grant with a
      // different role cleanly replaces the prior role (mirrors slack-app).
      try {
        await this._executeCommandAsync('nsf-share-record ' + recordUid + ' -e ' + userEmail + ' -a revoke --force', 15);
      } catch (e) {
        log.debug('NSF record revoke before grant skipped or failed, proceeding', e.message);
      }

      let command = 'nsf-share-record ' + recordUid + ' -e ' + userEmail + ' -a grant -r ' + role;

      let expiresAtStr = 'Never (Permanent)';
      const isPermanent = durationSeconds === null;
      if (!isPermanent && durationSeconds) {
        const expireIn = this._formatDuration(durationSeconds);
        command += ' --expire-in ' + expireIn;
        const expiresAt = new Date(Date.now() + durationSeconds * 1000);
        expiresAtStr = expiresAt.toISOString();
      }

      command += ' --force';

      log.info('Executing NSF record grant command', { command });
      const result = await this._executeCommandAsync(command, 30);
      log.info('NSF record grant result', { status: result?.status, message: result?.message, error: result?.error });

      if (result?.status === 'success') {
        const message = result?.message || '';
        const invitationSent = this._isInvitationSent(message);
        return {
          success: true,
          expiresAt: expiresAtStr,
          permission: role,
          nsfRole: role,
          isNsf: true,
          duration: isPermanent ? 'permanent' : 'temporary',
          invitationSent: invitationSent,
        };
      }

      const errorMsg = this._formatError(result?.message) || '';
      const errorField = result?.error || '';
      const combined = (errorMsg + ' ' + errorField).trim();

      if (this._isInvitationSent(combined)) {
        log.info('NSF record share invitation sent (detected in error response)', { recordUid, userEmail });
        return {
          success: true,
          expiresAt: 'Pending Invitation',
          permission: role,
          nsfRole: role,
          isNsf: true,
          duration: 'permanent',
          invitationSent: true,
        };
      }

      return { success: false, error: errorMsg || errorField || 'Failed to share Nested Share record' };
    } catch (error) {
      return { success: false, error: 'Error granting Nested Share record access: ' + error.message };
    }
  }

  /**
   * Grant a user access to a Nested Share Folder (NSF) via `nsf-share-folder`.
   * @param {string} folderUid - NSF folder UID
   * @param {string} userEmail - Recipient email
   * @param {string} role - NSF role
   * @param {number|null} durationSeconds - Access duration in seconds; null = permanent
   * @returns {Promise<object>} Result shaped like grantFolderAccess
   */
  async grantNsfFolderAccess(folderUid, userEmail, role, durationSeconds = 86400) {
    try {
      if (!isValidNsfRole(role)) {
        return { success: false, error: `Invalid Nested Share role: ${role}` };
      }

      try { await this._executeCommandAsync('sync-down', 15); } catch (e) { log.warn('sync-down before nsf-share-folder failed, proceeding', e.message); }

      let command = 'nsf-share-folder ' + folderUid + ' -a grant -e ' + userEmail + ' -r ' + role;

      let expiresAtStr = 'Never (Permanent)';
      const isPermanent = durationSeconds === null;
      if (!isPermanent && durationSeconds) {
        const expireIn = this._formatDuration(durationSeconds);
        command += ' --expire-in ' + expireIn;
        const expiresAt = new Date(Date.now() + durationSeconds * 1000);
        expiresAtStr = expiresAt.toISOString();
      }

      log.info('Executing NSF folder grant command', { command });
      const result = await this._executeCommandAsync(command, 30);
      log.info('NSF folder grant result', { status: result?.status, message: result?.message, error: result?.error });

      if (result?.status === 'success') {
        const message = result?.message || '';
        const invitationSent = this._isInvitationSent(message);
        return {
          success: true,
          expiresAt: expiresAtStr,
          permission: role,
          nsfRole: role,
          isNsf: true,
          duration: isPermanent ? 'permanent' : 'temporary',
          invitationSent: invitationSent,
        };
      }

      const errorMsg = this._formatError(result?.message) || '';
      const errorField = result?.error || '';
      const combined = (errorMsg + ' ' + errorField).trim();

      if (this._isInvitationSent(combined)) {
        log.info('NSF folder share invitation sent (detected in error response)', { folderUid, userEmail });
        return {
          success: true,
          expiresAt: 'Pending Invitation',
          permission: role,
          nsfRole: role,
          isNsf: true,
          duration: 'permanent',
          invitationSent: true,
        };
      }

      return { success: false, error: errorMsg || errorField || 'Failed to share Nested Share folder' };
    } catch (error) {
      return { success: false, error: 'Error granting Nested Share folder access: ' + error.message };
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

  /**
   * Shared folders visible to the vault user from `share-report -f` (for create-secret folder picker).
   * Rows are filtered to those where "Shared To" matches userEmail (case-insensitive).
   * @param {string} userEmail - Requester work email (Graph)
   * @returns {Promise<{ choiceSetChoices: Array<{title:string,value:string}>, error: string|null, noSharedFoldersForUser: boolean }>}
   */
  async getSharedFolderChoicesForEmail(userEmail) {
    const defaultChoice = { title: 'My vault (default)', value: '_default_' };
    const loadErrorResult = {
      choiceSetChoices: [],
      error: 'Unable to load shared folder list. Please try again.',
      noSharedFoldersForUser: false,
    };
    try {
      if (!userEmail || !String(userEmail).trim()) {
        log.warn('getSharedFolderChoicesForEmail: email not resolved');
        return loadErrorResult;
      }
      const command = 'share-report -f --format=json';
      const result = await this._executeCommandAsync(command, 45);
      if (!result || result.status !== 'success') {
        const errMsg = this._formatError(result?.message) || 'Could not load shared folders';
        log.warn('getSharedFolderChoicesForEmail failed', errMsg);
        return loadErrorResult;
      }

      let rows = result.data;
      if (typeof rows === 'string') {
        try {
          rows = JSON.parse(rows);
        } catch {
          rows = [];
        }
      }
      if (!Array.isArray(rows)) {
        rows = [];
      }

      const emailLower = String(userEmail).trim().toLowerCase();
      const seen = new Set();
      const out = [];

      for (const row of rows) {
        const sharedTo = String(row['Shared To'] || row.shared_to || row.SharedTo || '').trim().toLowerCase();
        if (!sharedTo || sharedTo !== emailLower) {
          continue;
        }
        const uid = row['Folder UID'] || row.folder_uid || row.folderUid;
        if (!uid || seen.has(uid)) {
          continue;
        }
        seen.add(uid);
        const name = row['Folder Name'] || row.folder_name || uid;
        const path = row['Folder Path'] || row['folder_path'] || name;
        const folderType = row['Type'] || row.type || row.folder_type || '';
        const label = path && path !== name ? `${name} (${path})` : String(name);
        out.push({
          title: label.length > 120 ? label.slice(0, 117) + '...' : label,
          value: uid,
          isNsf: isNsfType(folderType),
        });
      }

      out.sort((a, b) => a.title.localeCompare(b.title, undefined, { sensitivity: 'base' }));
      if (out.length === 0) {
        return {
          choiceSetChoices: [],
          error: null,
          noSharedFoldersForUser: true,
        };
      }
      return {
        choiceSetChoices: out,
        error: null,
        noSharedFoldersForUser: false,
      };
    } catch (error) {
      log.error('getSharedFolderChoicesForEmail exception', error.message);
      return loadErrorResult;
    }
  }

  /**
   * Get subfolders inside a shared folder using `tree -s -v <folderUid>`.
   * Returns a flat list of all subfolders (including nested) for the dropdown.
   * @param {string} folderUid - The shared folder UID
   * @returns {Promise<{ choices: Array<{title:string,value:string}>, error: string|null }>}
   */
  async getSubfoldersForSharedFolder(folderUid) {
    try {
      if (!folderUid || !String(folderUid).trim()) {
        return { choices: [], error: null };
      }
      const command = 'tree -s -v ' + shellEscapeSearch(String(folderUid).trim());
      const result = await this._executeCommandAsync(command, 20);
      if (!result || result.status !== 'success') {
        const errMsg = this._formatError(result?.message) || 'Could not load subfolders';
        log.warn('getSubfoldersForSharedFolder failed', errMsg);
        return { choices: [], error: 'Unable to load subfolders. Please try again.' };
      }

      const tree = result.data?.tree;
      if (!Array.isArray(tree) || tree.length === 0) {
        return { choices: [], error: null };
      }

      const choices = tree
        .filter((item) => item.type === 'folder' && item.uid)
        .map((item) => ({
          title: item.path || item.name || item.uid,
          value: item.uid,
        }));

      return { choices, error: null };
    } catch (error) {
      log.error('getSubfoldersForSharedFolder exception', error.message);
      return { choices: [], error: 'Unable to load subfolders. Please try again.' };
    }
  }

  async createRecord({
    title,
    login,
    password,
    url,
    notes,
    generatePassword = false,
    selfDestructDuration = null,
    folderUid = null,
    skipUidLookup = false,
  }) {
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

      if (folderUid && String(folderUid).trim() && String(folderUid).trim() !== '_default_') {
        commandParts.push('--folder ' + shellEscapeSearch(String(folderUid).trim()));
      }
      
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
        const classified = this._classifyCreateRecordError(result);
        return {
          success: false,
          error: classified.friendly,
          errorCode: classified.code,
          errorTitle: classified.title,
          rawError: classified.raw,
          httpStatus: result?.httpStatus,
        };
      }

      const recordUid = this._extractRecordUidFromCreateResult(result);
      if (recordUid) {
        log.info('Record created with UID from response', { recordUid, title });
        return {
          success: true,
          recordUid,
          title,
          generatedPassword: generatePassword || password === '$GEN',
          isNewlyCreated: true,
          selfDestructUrl: selfDestructDuration ? this._extractUrlFromCreateResult(result) : null,
        };
      }
      
      if (skipUidLookup) {
        log.warn('Record created but UID not returned in response; skipping post-create UID lookup', { title });
        return {
          success: true,
          recordUid: null,
          title,
          generatedPassword: generatePassword || password === '$GEN',
          isNewlyCreated: true,
          uidUnavailable: true,
          selfDestructUrl: selfDestructDuration ? this._extractUrlFromCreateResult(result) : null,
        };
      }

      log.debug('Record created successfully, searching for UID...');

      const maxRetries = 3;
      const waitTimes = [2000, 3000, 4000];
      
      for (let attempt = 0; attempt < maxRetries; attempt++) {
        await new Promise(resolve => setTimeout(resolve, waitTimes[attempt]));
        
        try {
          const searchCommand = 'search ' + shellEscape(title) + ' --format=json';
          const searchResult = await this._executeCommandAsync(searchCommand, 10);
          
          if (searchResult && searchResult.status === 'success' && searchResult.data) {
            const data = Array.isArray(searchResult.data) ? searchResult.data : [];
            
            if (data.length > 0) {
              const matchingRecord = data.find(r => r.title === title) || data[0];
              const uid = matchingRecord.uid || matchingRecord.record_uid;
              
              if (uid) {
                log.info('Found created record UID after attempt ' + (attempt + 1), { recordUid: uid, title });
                return {
                  success: true,
                  recordUid: uid,
                  title,
                  generatedPassword: generatePassword || password === '$GEN',
                  isNewlyCreated: true,
                  selfDestructUrl: selfDestructDuration ? this._extractUrlFromCreateResult(result) : null,
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
        success: true,
        recordUid: null,
        title,
        generatedPassword: generatePassword || password === '$GEN',
        isNewlyCreated: true,
        uidUnavailable: true,
        selfDestructUrl: selfDestructDuration ? this._extractUrlFromCreateResult(result) : null,
      };
    } catch (error) {
      log.error('Exception in createRecord', error);
      return { success: false, error: 'Error creating record: ' + error.message };
    }
  }

  /**
   * Create a record inside a Nested Share Folder (NSF) via `nsf-record-add`.
   * Mirrors createRecord but targets the NSF (v3) folder system.
   * @returns {Promise<object>} Result shaped like createRecord
   */
  async createNsfRecord({
    title,
    login,
    password,
    url,
    notes,
    generatePassword = false,
    folderUid = null,
  }) {
    try {
      if (!title || !title.trim()) {
        return { success: false, error: 'Title is required' };
      }

      const commandParts = ['nsf-record-add'];
      commandParts.push('--record-type login');
      commandParts.push('--title ' + shellEscape(title));

      if (folderUid && String(folderUid).trim() && String(folderUid).trim() !== '_default_') {
        commandParts.push('--folder ' + shellEscapeSearch(String(folderUid).trim()));
      }

      if (notes && notes.trim()) {
        const notesForCli = notes.replace(/\n/g, '\\n');
        commandParts.push('--notes ' + shellEscape(notesForCli));
      }

      if (login && login.trim()) {
        commandParts.push('login=' + shellEscape(login));
      }

      if (password && password.trim()) {
        if (password === '$GEN') {
          commandParts.push('password=$GEN');
        } else {
          commandParts.push('password=' + shellEscape(password));
        }
      } else if (generatePassword) {
        commandParts.push('password=$GEN');
      }

      if (url && url.trim()) {
        commandParts.push('url=' + shellEscape(url));
      }

      const command = commandParts.join(' ');
      log.debug('Creating NSF record with command', command.replace(/password=[^\s]+/, 'password=***'));

      const result = await this._executeCommandAsync(command, 20);

      if (!result || result.status !== 'success') {
        const classified = this._classifyCreateRecordError(result);
        return {
          success: false,
          error: classified.friendly,
          errorCode: classified.code,
          errorTitle: classified.title,
          rawError: classified.raw,
          httpStatus: result?.httpStatus,
        };
      }

      const recordUid = this._extractRecordUidFromCreateResult(result);
      if (recordUid) {
        log.info('NSF record created with UID from response', { recordUid, title });
      } else {
        // NSF records are not indexed by the classic `search` command, so the UID
        // may not be resolvable post-create. The record still exists in the vault.
        log.warn('NSF record created but UID not returned in response', { title });
      }

      return {
        success: true,
        recordUid,
        title,
        generatedPassword: generatePassword || password === '$GEN',
        isNewlyCreated: true,
        isNsf: true,
        uidUnavailable: !recordUid,
      };
    } catch (error) {
      log.error('Exception in createNsfRecord', error);
      return { success: false, error: 'Error creating Nested Share record: ' + error.message };
    }
  }

  _extractRecordUidFromCreateResult(result) {
    if (!result) return null;

    const directContainers = [result.data, result];
    const directKeys = [
      'record_uid',
      'recordUid',
      'recordUID',
      'uid',
      'UID',
      'record_id',
      'recordId',
    ];

    for (const container of directContainers) {
      if (!container || typeof container !== 'object') continue;
      for (const key of directKeys) {
        const value = container[key];
        if (typeof value === 'string' && this._looksLikeKeeperUid(value)) {
          return value;
        }
      }
    }

    const strings = [];
    const collectStrings = (value, depth = 0) => {
      if (value == null || depth > 5) return;
      if (typeof value === 'string') {
        strings.push(value);
        return;
      }
      if (Array.isArray(value)) {
        value.forEach((item) => collectStrings(item, depth + 1));
        return;
      }
      if (typeof value === 'object') {
        Object.values(value).forEach((item) => collectStrings(item, depth + 1));
      }
    };

    collectStrings(result.data);
    collectStrings(result.message);
    collectStrings(result.error);

    for (const text of strings) {
      const match = text.match(/\b[A-Za-z0-9_-]{20,24}\b/);
      if (match && this._looksLikeKeeperUid(match[0])) {
        return match[0];
      }
    }

    return null;
  }

  _extractUrlFromCreateResult(result) {
    if (!result) return null;

    const directContainers = [result.data, result];
    const directKeys = ['url', 'share_url', 'shareUrl', 'shareURL', 'self_destruct_url', 'selfDestructUrl', 'link'];

    for (const container of directContainers) {
      if (!container || typeof container !== 'object') continue;
      for (const key of directKeys) {
        const value = container[key];
        if (typeof value === 'string' && /^https?:\/\//i.test(value)) {
          return value;
        }
      }
    }

    const strings = [];
    const collectStrings = (value, depth = 0) => {
      if (value == null || depth > 5) return;
      if (typeof value === 'string') {
        strings.push(value);
        return;
      }
      if (Array.isArray(value)) {
        value.forEach((item) => collectStrings(item, depth + 1));
        return;
      }
      if (typeof value === 'object') {
        Object.values(value).forEach((item) => collectStrings(item, depth + 1));
      }
    };

    collectStrings(result.data);
    collectStrings(result.message);
    collectStrings(result.error);

    for (const text of strings) {
      const match = text.match(/https?:\/\/[^\s<>"')\]]+/i);
      if (match) {
        return match[0];
      }
    }

    return null;
  }

  _looksLikeKeeperUid(value) {
    return typeof value === 'string' && /^[A-Za-z0-9_-]{20,24}$/.test(value);
  }

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
      log.info('Approving EPM request', { approvalUid });
      
      const result = await this._executeCommandAsync(command, 30);
      
      log.debug('approvePedmRequest result', { status: result?.status });
      
      if (result?.status === 'success') {
        // Success! Request was approved
        log.info('EPM request approved successfully', { approvalUid });
        return { success: true };
      } else {
        // Check error message for "already processed" indicators
        const errorMsg = this._formatError(result?.message) || '';
        const errorField = result?.error || '';
        const combinedError = (errorMsg + ' ' + errorField).toLowerCase();
        
        // Match Slack's approach: only these specific error messages indicate already processed
        if (combinedError.includes('does not exist or cannot be modified') || 
            combinedError.includes('approval request does not exist') ||
            combinedError.includes('does not exist')) {
          log.info('EPM request already processed (detected from error)', { approvalUid, error: combinedError });
          return { 
            success: false, 
            error: 'This request has already been approved or denied.',
            already_processed: true 
          };
        }
        
        return { success: false, error: errorMsg || 'Failed to approve EPM request' };
      }
    } catch (error) {
      log.error('approvePedmRequest exception:', error.message);
      return { success: false, error: error.message };
    }
  }

  async denyPedmRequest(approvalUid) {
    try {
      const command = 'epm approval action --deny ' + approvalUid;
      log.info('Denying EPM request', { approvalUid });
      
      const result = await this._executeCommandAsync(command, 30);
      
      log.debug('denyPedmRequest result', { status: result?.status });
      
      if (result?.status === 'success') {
        // Success! Request was denied
        log.info('EPM request denied successfully', { approvalUid });
        return { success: true };
      } else {
        // Check error message for "already processed" indicators
        const errorMsg = this._formatError(result?.message) || '';
        const errorField = result?.error || '';
        const combinedError = (errorMsg + ' ' + errorField).toLowerCase();
        
        // Match Slack's approach: only these specific error messages indicate already processed
        if (combinedError.includes('does not exist or cannot be modified') || 
            combinedError.includes('approval request does not exist') ||
            combinedError.includes('does not exist')) {
          log.info('EPM request already processed (detected from error)', { approvalUid, error: combinedError });
          return { 
            success: false, 
            error: 'This request has already been approved or denied.',
            already_processed: true 
          };
        }
        
        return { success: false, error: errorMsg || 'Failed to deny EPM request' };
      }
    } catch (error) {
      log.error('denyPedmRequest exception:', error.message);
      return { success: false, error: error.message };
    }
  }

  async getPendingDeviceApprovals() {
    try {
      const command = 'device-approve --reload --format=json';
      const result = await this._executeCommandAsync(command, 30);
      
      log.debug('getPendingDeviceApprovals raw result', { 
        status: result?.status, 
        hasData: !!result?.data,
        dataType: Array.isArray(result?.data) ? 'array' : typeof result?.data
      });
      
      if (!result || result.status !== 'success') {
        log.debug('getPendingDeviceApprovals: no success status', { result });
        return [];
      }
      
      const data = result.data;
      if (!data || !Array.isArray(data)) {
        log.debug('getPendingDeviceApprovals: data is not an array', { data });
        return [];
      }
      
      log.debug('getPendingDeviceApprovals found devices', { 
        count: data.length,
        sample: data.length > 0 ? JSON.stringify(data[0]).substring(0, 200) : 'none'
      });
      
      return data;
    } catch (error) {
      log.error('Error getting device approvals', error.message);
      return [];
    }
  }

  async approveDevice(deviceId) {
    try {
      const pendingDevices = await this.getPendingDeviceApprovals();

      log.info('Checking pending devices', { 
        deviceId, 
        pendingCount: pendingDevices.length,
        pendingIds: pendingDevices.map(d => d.device_id || d.deviceId || d.id).slice(0, 10)
      });
      
      const isPending = pendingDevices.some(d => {
        const pendingId = d.device_id || d.deviceId || d.id;
        const matches = pendingId === deviceId;
        if (matches) {
          log.debug('Found matching pending device', { pendingId, deviceId });
        }
        return matches;
      });
      
      if (!isPending) {
        log.info('Device not in pending list - already processed elsewhere', { deviceId, pendingDevices: pendingDevices.slice(0, 3) });
        return { 
          success: false, 
          already_processed: true, 
          error: 'This request has already been approved or denied from another platform.' 
        };
      }
      
      // Device is pending, proceed with approval
      const command = 'device-approve --approve ' + deviceId;
      const result = await this._executeCommandAsync(command, 30);
      
      log.debug('approveDevice result', { status: result?.status });
      
      if (result?.status === 'success') {
        // Also check message and data for "no pending" indicators
        const message = result.message || '';
        const dataStr = typeof result.data === 'string' ? result.data : JSON.stringify(result.data || '');
        const combinedText = (message + ' ' + dataStr).toLowerCase();
        
        // Check if this is an "already processed" response
        // Note: Don't check for generic 'no output' - "Command executed successfully but produced no output" is a valid success
        if (combinedText.includes('no pending') || 
            combinedText.includes('there are no pending') ||
            combinedText.includes('already processed') ||
            combinedText.includes('already approved') ||
            combinedText.includes('already denied')) {
          log.info('Device already processed detected in success response', deviceId);
          return { success: false, already_processed: true, error: 'This request has already been approved or denied from another platform.' };
        }
        
        // Success! Device was approved
        log.info('Device approved successfully', { deviceId });
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
      const pendingDevices = await this.getPendingDeviceApprovals();
      
      log.info('Checking pending devices for deny', { 
        deviceId, 
        pendingCount: pendingDevices.length,
        pendingIds: pendingDevices.map(d => d.device_id || d.deviceId || d.id).slice(0, 10)
      });
      
      const isPending = pendingDevices.some(d => {
        const pendingId = d.device_id || d.deviceId || d.id;
        return pendingId === deviceId;
      });
      
      if (!isPending) {
        log.info('Device not in pending list - already processed elsewhere', { deviceId });
        return { 
          success: false, 
          already_processed: true, 
          error: 'This request has already been approved or denied from another platform.' 
        };
      }
      
      // Device is pending, proceed with denial
      const command = 'device-approve --deny ' + deviceId;
      const result = await this._executeCommandAsync(command, 30);
      
      log.debug('denyDevice result', { status: result?.status });
      
      if (result?.status === 'success') {
        // Success! Device was denied
        log.info('Device denied successfully', { deviceId });
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
      log.debug('Executing command', command.replace(/password=[^\s]+/, 'password=***'));
      
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
        let recordCategory = '';
        
        if (item.details) {
          const parts = item.details.split(', ');
          for (const part of parts) {
            if (part.startsWith('Type: ')) {
              recordType = part.replace('Type: ', '').trim();
            } else if (part.startsWith('Description: ')) {
              notes = part.replace('Description: ', '').trim();
            } else if (part.startsWith('Record Category: ')) {
              recordCategory = part.replace('Record Category: ', '').trim();
            }
          }
        }
        
        results.push({
          uid: item.uid,
          title: item.name,
          recordType: recordType,
          notes: notes,
          // NSF records keep their normal type (e.g. "login"); the reliable
          // signal is "Record Category: Nested". Keep the type check as a fallback.
          isNsf: isNsfRecordCategory(recordCategory) || isNsfType(item.type),
        });
      } else {
        results.push({
          uid: item.uid,
          name: item.name,
          folderType: item.type || 'shared_folder',
          isNsf: isNsfType(item.type),
        });
      }
    }
    
    return results;
  }

  _classifyCreateRecordError(result) {
    const raw = String(result?.error || result?.message || '').trim();
    const lower = raw.toLowerCase();

    if (lower.includes('restricted by your enterprise role policy')) {
      const m = raw.match(/record type "([^"]+)" is restricted/i);
      const recordType = m ? m[1] : 'login';
      return {
        code: 'POLICY_RECORD_TYPE_RESTRICTED',
        title: 'Record type not permitted',
        friendly:
          `Your administrator has restricted creation of "${recordType}" records ` +
          `via the Keeper Admin Console role policy. Please contact your Keeper administrator.`,
        raw,
      };
    }

    if (
      lower.includes('password must be at least') ||
      lower.includes('password must contain at least')
    ) {
      const cleaned = raw
        .split('\n')
        .map((s) => s.trim())
        .filter((line) => line && !/^use --force/i.test(line));
      const bulleted = cleaned.length
        ? cleaned.map((line) => '• ' + line).join('\n')
        : raw;
      return {
        code: 'POLICY_PASSWORD_COMPLEXITY',
        title: "Password does not meet your organization's policy",
        friendly: bulleted,
        raw,
      };
    }

    return {
      code: 'GENERIC',
      title: 'Record creation failed',
      friendly: this._formatError(raw),
      raw,
    };
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

// Export instance
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
    isPamUserFolder: async () => false,
    grantRecordAccess: async () => ({ success: false, error: 'Client not initialized' }),
    grantFolderAccess: async () => ({ success: false, error: 'Client not initialized' }),
    grantNsfRecordAccess: async () => ({ success: false, error: 'Client not initialized' }),
    grantNsfFolderAccess: async () => ({ success: false, error: 'Client not initialized' }),
    createOneTimeShare: async () => ({ success: false, error: 'Client not initialized' }),
    createRecord: async () => ({ success: false, error: 'Client not initialized' }),
    createNsfRecord: async () => ({ success: false, error: 'Client not initialized' }),
    getSharedFolderChoicesForEmail: async () => ({
      choiceSetChoices: [],
      error: 'Unable to load shared folder list. Please try again.',
      noSharedFoldersForUser: false,
    }),
    getSubfoldersForSharedFolder: async () => ({ choices: [], error: 'Client not initialized' }),
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
