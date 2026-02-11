/**
 * Adaptive Card builders for approval requests
 * 
 * This file re-exports all card builders from modular files.
 * Cards are displayed in the approvals channel when users request access
 * to records, folders, or one-time shares.
 */

// Import constants
const { 
  RECORD_PERMISSIONS, 
  FOLDER_PERMISSIONS, 
  DURATION_OPTIONS,
  SHARE_DURATION_OPTIONS,
} = require('./constants');

// Import helper functions
const {
  formatPermissionLabel,
  formatFolderPermissionLabel,
} = require('./cardHelpers');

// Import record cards
const {
  buildRecordApprovalCard,
  buildRecordSearchResultsCard,
  buildRecordApprovalCardWithStatus,
  buildRecordConfirmationCard,
  buildRecordCreationCard,
  buildRecordCreatedCard,
} = require('./approval/recordRequestCards');

// Import folder cards
const {
  buildFolderApprovalCard,
  buildFolderSearchResultsCard,
  buildFolderApprovalCardWithStatus,
  buildFolderConfirmationCard,
} = require('./approval/folderRequestCards');

// Import share cards
const {
  buildShareSearchResultsCard,
  buildOneTimeShareApprovalCard,
  buildOneTimeShareApprovalCardWithStatus,
} = require('./approval/oneTimeShareCards');

// Aliases for backward compatibility (used in commandHandler.js)
const createRecordApprovalCard = buildRecordApprovalCard;
const createFolderApprovalCard = buildFolderApprovalCard;

module.exports = {
  // Record cards
  buildRecordApprovalCard,
  buildRecordApprovalCardWithStatus,
  buildRecordConfirmationCard,
  buildRecordSearchResultsCard,
  buildRecordCreationCard,
  buildRecordCreatedCard,
  
  // Folder cards
  buildFolderApprovalCard,
  buildFolderApprovalCardWithStatus,
  buildFolderConfirmationCard,
  buildFolderSearchResultsCard,
  
  // Share cards
  buildShareSearchResultsCard,
  buildOneTimeShareApprovalCard,
  buildOneTimeShareApprovalCardWithStatus,
  
  // Aliases (backward compatibility)
  createRecordApprovalCard,
  createFolderApprovalCard,
  
  // Constants
  RECORD_PERMISSIONS,
  FOLDER_PERMISSIONS,
  DURATION_OPTIONS,
  SHARE_DURATION_OPTIONS,
  
  // Helper functions
  formatPermissionLabel,
  formatFolderPermissionLabel,
};
