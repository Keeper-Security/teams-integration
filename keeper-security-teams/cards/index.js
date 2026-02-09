/**
 * Cards index - export all Adaptive Card builders
 */

const approvalCard = require('./approvalCard');
const resultCard = require('./resultCard');
const pedmCard = require('./pedmCard');
const deviceCard = require('./deviceCard');
const createRecordCard = require('./createRecordCard');

module.exports = {
  // Approval request cards
  buildRecordApprovalCard: approvalCard.buildRecordApprovalCard,
  buildRecordApprovalCardWithStatus: approvalCard.buildRecordApprovalCardWithStatus,
  buildRecordConfirmationCard: approvalCard.buildRecordConfirmationCard,
  buildRecordSearchResultsCard: approvalCard.buildRecordSearchResultsCard,
  buildFolderApprovalCard: approvalCard.buildFolderApprovalCard,
  buildFolderApprovalCardWithStatus: approvalCard.buildFolderApprovalCardWithStatus,
  buildFolderConfirmationCard: approvalCard.buildFolderConfirmationCard,
  buildFolderSearchResultsCard: approvalCard.buildFolderSearchResultsCard,
  buildShareSearchResultsCard: approvalCard.buildShareSearchResultsCard,
  buildOneTimeShareApprovalCard: approvalCard.buildOneTimeShareApprovalCard,
  buildOneTimeShareApprovalCardWithStatus: approvalCard.buildOneTimeShareApprovalCardWithStatus,
  createRecordApprovalCard: approvalCard.createRecordApprovalCard,
  createFolderApprovalCard: approvalCard.createFolderApprovalCard,
  
  // Result/notification cards
  buildApprovalResultCard: resultCard.buildApprovalResultCard,
  buildShareResultCard: resultCard.buildShareResultCard,
  buildSearchResultsCard: resultCard.buildSearchResultsCard,
  buildHelpCard: resultCard.buildHelpCard,
  buildErrorCard: resultCard.buildErrorCard,
  buildApprovedMessageCard: resultCard.buildApprovedMessageCard,
  buildDeniedMessageCard: resultCard.buildDeniedMessageCard,
  buildRequesterNotificationCard: resultCard.buildRequesterNotificationCard,
  createShareResultCard: resultCard.createShareResultCard,
  createSearchResultsCard: resultCard.createSearchResultsCard,
  createHelpCard: resultCard.createHelpCard,
  
  // PEDM cards
  buildPedmApprovalCard: pedmCard.buildPedmApprovalCard,
  buildPedmApprovedCard: pedmCard.buildPedmApprovedCard,
  buildPedmDeniedCard: pedmCard.buildPedmDeniedCard,
  
  // Device approval cards
  buildDeviceApprovalCard: deviceCard.buildDeviceApprovalCard,
  buildDeviceApprovedCard: deviceCard.buildDeviceApprovedCard,
  buildDeviceDeniedCard: deviceCard.buildDeviceDeniedCard,
  
  // Constants
  RECORD_PERMISSIONS: approvalCard.RECORD_PERMISSIONS,
  FOLDER_PERMISSIONS: approvalCard.FOLDER_PERMISSIONS,
  DURATION_OPTIONS: approvalCard.DURATION_OPTIONS,
  SHARE_DURATION_OPTIONS: approvalCard.SHARE_DURATION_OPTIONS,
  
  // Create record cards
  buildCreateRecordModal: createRecordCard.buildCreateRecordModal,
  buildCreateRecordSuccessResponse: createRecordCard.buildCreateRecordSuccessResponse,
  
  // Helpers
  formatPermission: resultCard.formatPermission,
  formatDate: resultCard.formatDate,
  getRecordIcon: resultCard.getRecordIcon,
};
