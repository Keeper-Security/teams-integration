/**
 * Helper functions for building Adaptive Cards
 * Shared utilities used across record, folder, and share cards
 */

const { RECORD_PERMISSIONS, FOLDER_PERMISSIONS } = require('./constants');
const { sanitizeHyperlinks } = require('../utils/helpers');

/**
 * Build common header section for search results cards
 * @param {string} title - Card title (e.g., "Record Access Request")
 * @param {string} requesterName - Name of the requester
 * @param {string} approvalId - Approval request ID
 * @param {string} justification - Request justification
 * @returns {Array} Array of card body elements
 */
function buildSearchCardHeader(title, requesterName, approvalId, justification) {
  // Sanitize justification to prevent URL injection
  const safeJustification = sanitizeHyperlinks(justification) || 'No justification provided';
  
  return [
    {
      type: 'TextBlock',
      text: title,
      weight: 'Bolder',
      size: 'ExtraLarge',
    },
    {
      type: 'ColumnSet',
      columns: [
        {
          type: 'Column',
          width: 'stretch',
          items: [
            {
              type: 'TextBlock',
              text: 'Requester:',
              weight: 'Bolder',
              size: 'Medium',
            },
            {
              type: 'TextBlock',
              text: requesterName || 'Unknown',
              color: 'Warning',
              size: 'Medium',
            },
          ],
        },
        {
          type: 'Column',
          width: 'stretch',
          items: [
            {
              type: 'TextBlock',
              text: 'Request ID:',
              weight: 'Bolder',
              size: 'Medium',
            },
            {
              type: 'TextBlock',
              text: approvalId || 'N/A',
              color: 'Warning',
              size: 'Medium',
            },
            {
              type: 'TextBlock',
              text: 'Justification:',
              weight: 'Bolder',
              size: 'Medium',
              spacing: 'Medium',
            },
            {
              type: 'TextBlock',
              text: safeJustification,
              wrap: true,
              size: 'Medium',
            },
          ],
        },
      ],
    },
  ];
}

/**
 * Header for the self-service /keeper-create-secret form (not the admin approval create flow).
 * Internal request IDs stay only on Action.Execute payloads; users see who is creating the record.
 */
function buildCreateSecretHeader(userName, userEmail) {
  return [
    {
      type: 'TextBlock',
      text: 'Create a new secret',
      weight: 'Bolder',
      size: 'ExtraLarge',
    },
    {
      type: 'TextBlock',
      text: 'Add a login record to your Keeper vault. Leave password empty to auto-generate.',
      wrap: true,
      isSubtle: true,
      spacing: 'Small',
    },
    {
      type: 'FactSet',
      facts: [
        { title: 'User', value: userName || 'Unknown' },
        { title: 'Email', value: userEmail || '—' },
      ],
    },
  ];
}

/**
 * Build "no results" section for search cards
 * @param {string} searchQuery - The search query that returned no results
 * @param {string} itemType - 'record', 'folder', or 'share'
 * @returns {Array} Array of card body elements for no results state
 */
function buildNoResultsSection(searchQuery, itemType) {
  const itemLabel = itemType === 'folder' ? 'folders' : 'records';
  return [
    {
      type: 'Container',
      style: 'attention',
      spacing: 'Medium',
      items: [
        {
          type: 'TextBlock',
          text: `No ${itemLabel} found for "${searchQuery}"`,
          wrap: true,
          weight: 'Bolder',
        },
      ],
    },
    {
      type: 'TextBlock',
      text: 'Try a different search term:',
      wrap: true,
      spacing: 'Medium',
    },
    {
      type: 'Input.Text',
      id: 'searchQuery',
      placeholder: `Enter ${itemType === 'folder' ? 'folder' : 'record'} name or UID...`,
      value: searchQuery || '',
    },
  ];
}

/**
 * Build "found items" header section
 * @param {number} count - Number of items found
 * @param {string} itemType - 'record' or 'folder'
 * @param {Object} singleItem - The item if only one found (optional)
 * @returns {Array} Array of card body elements
 */
function buildFoundItemsHeader(count, itemType, singleItem = null) {
  const itemLabel = itemType === 'folder' ? 'Folder' : 'Record';
  const itemsLabel = itemType === 'folder' ? 'Folders' : 'Records';
  
  if (count === 1 && singleItem) {
    return [
      {
        type: 'Container',
        style: 'good',
        spacing: 'Medium',
        items: [
          {
            type: 'TextBlock',
            text: `${itemLabel} Found: ${singleItem.title || singleItem.name}`,
            wrap: true,
            weight: 'Bolder',
          },
          {
            type: 'TextBlock',
            text: `UID: ${singleItem.uid}`,
            size: 'Small',
            isSubtle: true,
          },
        ],
      },
    ];
  }
  
  return [
    {
      type: 'Container',
      style: 'good',
      spacing: 'Medium',
      items: [
        {
          type: 'TextBlock',
          text: `${count} ${itemsLabel} Found`,
          wrap: true,
          weight: 'Bolder',
        },
        {
          type: 'TextBlock',
          text: `Select the correct ${itemLabel.toLowerCase()} from the list below:`,
          size: 'Small',
          isSubtle: true,
        },
      ],
    },
  ];
}

/**
 * Format permission label for display
 * @param {string} permission - Permission value
 * @returns {string} Formatted permission label
 */
function formatPermissionLabel(permission) {
  const found = RECORD_PERMISSIONS.find(p => p.value === permission);
  return found ? found.title : permission;
}

/**
 * Format folder permission label for display
 * @param {string} permission - Permission value
 * @returns {string} Formatted permission label
 */
function formatFolderPermissionLabel(permission) {
  const found = FOLDER_PERMISSIONS.find(p => p.value === permission);
  return found ? found.title : permission;
}

/**
 * Get current timestamp formatted for display
 * Returns format: "YYYY-MM-DD HH:MM:SS"
 * @returns {string} Formatted timestamp
 */
function getCurrentTimestamp() {
  return new Date().toISOString().replace('T', ' ').substring(0, 19);
}

module.exports = {
  buildSearchCardHeader,
  buildCreateSecretHeader,
  buildNoResultsSection,
  buildFoundItemsHeader,
  formatPermissionLabel,
  formatFolderPermissionLabel,
  getCurrentTimestamp,
};
