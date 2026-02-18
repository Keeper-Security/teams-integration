/**
 * Microsoft Graph API Service
 * 
 * Service for fetching user information from Microsoft Graph API
 */

const axios = require('axios');
const { ManagedIdentityCredential, ClientSecretCredential } = require('@azure/identity');
const config = require('../config');
const { createLogger } = require('./logger');

const log = createLogger('GraphService');

class GraphService {
  constructor() {
    this.credential = null;
    this.tokenCache = new Map(); // Simple in-memory token cache
    this.tokenExpiry = new Map(); // Track token expiry
  }

  /**
   * Get Azure credential for Graph API calls
   * Uses Managed Identity in production, Client Secret in local dev
   * @returns {ManagedIdentityCredential|ClientSecretCredential}
   */
  getCredential() {
    if (!this.credential) {
      // Check if we're in local dev (has CLIENT_PASSWORD or CLIENT_SECRET) or production (Managed Identity)
      // Support both CLIENT_PASSWORD (from config) and CLIENT_SECRET (from env directly)
      const clientSecret = config.MicrosoftAppPassword || process.env.CLIENT_SECRET;
      
      if (clientSecret && config.MicrosoftAppTenantId) {
        // Local development: Use Client Secret
        log.debug('Using ClientSecretCredential for local development');
        this.credential = new ClientSecretCredential(
          config.MicrosoftAppTenantId,
          config.MicrosoftAppId,
          clientSecret
        );
      } else {
        // Production: Use Managed Identity
        log.debug('Using ManagedIdentityCredential for production');
        this.credential = new ManagedIdentityCredential({
          clientId: config.MicrosoftAppId,
        });
      }
    }
    return this.credential;
  }

  /**
   * Get access token for Microsoft Graph API
   * @param {string} tenantId - Optional tenant ID
   * @returns {Promise<string>} - Access token
   */
  async getAccessToken(tenantId = null) {
    const cacheKey = tenantId || 'default';
    const cachedToken = this.tokenCache.get(cacheKey);
    const expiry = this.tokenExpiry.get(cacheKey);
    
    // Return cached token if still valid (with 5 minute buffer)
    if (cachedToken && expiry && Date.now() < expiry - 300000) {
      log.debug('Using cached token');
      return cachedToken;
    }

    log.debug('Acquiring new access token...');
    try {
      const credential = this.getCredential();
      // Client credentials flow (ClientSecretCredential) MUST use .default scope
      // This will include all Application permissions granted to the app
      const scopes = ['https://graph.microsoft.com/.default'];
      const options = tenantId ? { tenantId } : {};
      
      const tokenResponse = await credential.getToken(scopes, options);

      // Decode token to see what roles/permissions we got (for debugging)
      try {
        const tokenParts = tokenResponse.token.split('.');
        if (tokenParts.length === 3) {
          const payload = JSON.parse(Buffer.from(tokenParts[1], 'base64').toString());
          // For application permissions, check 'roles' field
          // For delegated permissions, check 'scp' field
          const roles = payload.roles || [];
          const scopesInToken = payload.scp || [];
          log.debug('Token roles (application permissions)', roles.length > 0 ? roles : 'None');
          log.debug('Token scopes (delegated permissions)', scopesInToken.length > 0 ? scopesInToken : 'None');
          if (roles.length === 0 && scopesInToken.length === 0) {
            log.warn('Token has no roles or scopes!');
          }
        }
      } catch (decodeError) {
        // Ignore decode errors, just for debugging
        log.debug('Could not decode token for debugging');
      }

      // Cache the token (tokens typically expire in 1 hour)
      const expiresIn = (tokenResponse.expiresOnTimestamp - Date.now()) || 3600000;
      this.tokenCache.set(cacheKey, tokenResponse.token);
      this.tokenExpiry.set(cacheKey, Date.now() + expiresIn);

      log.debug('Successfully acquired access token');
      return tokenResponse.token;
    } catch (error) {
      log.error('Error getting access token', error.message);
      throw error;
    }
  }

  /**
   * Fetch user information from Microsoft Graph API
   * @param {string} userId - User ID (AAD Object ID or UPN)
   * @param {string} tenantId - Optional tenant ID
   * @returns {Promise<Object>} - User object with email and other properties
   */
  async getUser(userId, tenantId = null) {
    if (!userId) {
      throw new Error('User ID is required');
    }

    const cacheKey = tenantId || 'default';
    let retryCount = 0;
    const maxRetries = 1;

    while (retryCount <= maxRetries) {
      try {
        const token = await this.getAccessToken(tenantId);
        
        // Try to get user by object ID first
        let response;
        try {
          response = await axios.get(
            `https://graph.microsoft.com/v1.0/users/${userId}`,
            {
              headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json',
              },
            }
          );
          return response.data;
        } catch (error) {
          // If we get 403, it might be a stale token - clear cache and retry once
          if (error.response?.status === 403 && retryCount < maxRetries) {
            log.debug('Got 403 (Authorization denied), clearing token cache and retrying...');
            this.tokenCache.delete(cacheKey);
            this.tokenExpiry.delete(cacheKey);
            retryCount++;
            continue; // Retry with fresh token
          }
          
          // If userId is not an object ID, try as UPN
          if (error.response?.status === 404) {
            // If it looks like an email, try it as UPN
            if (userId.includes('@')) {
              response = await axios.get(
                `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}`,
                {
                  headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json',
                  },
                }
              );
              return response.data;
            } else {
              throw error;
            }
          } else {
            throw error;
          }
        }
      } catch (error) {
        // If we've exhausted retries or it's not a 403, throw the error
        log.error('Error fetching user', error.message);
        if (error.response) {
          log.error('Response status', error.response.status);
          log.error('Response data', error.response.data);
        }
        throw error;
      }
    }
  }

  /**
   * Get user email address from Microsoft Graph API
   * @param {string} userId - User ID (AAD Object ID)
   * @param {string} tenantId - Optional tenant ID
   * @returns {Promise<string|null>} - User email address or null if not found
   */
  async getUserEmail(userId, tenantId = null) {
    if (!userId) {
      return null;
    }

    try {
      const user = await this.getUser(userId, tenantId);
      // Prefer mail over userPrincipalName (mail is the primary email)
      return user.mail || user.userPrincipalName || null;
    } catch (error) {
      log.error('Error getting user email', error.message);
      return null;
    }
  }

  /**
   * Clear the token cache (useful when permissions are updated)
   * @param {string} tenantId - Optional tenant ID
   */
  clearTokenCache(tenantId = null) {
    const cacheKey = tenantId || 'default';
    this.tokenCache.delete(cacheKey);
    this.tokenExpiry.delete(cacheKey);
    log.debug('Token cache cleared for', cacheKey);
  }
}

// Export singleton instance
const graphService = new GraphService();

module.exports = graphService;
