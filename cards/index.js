/**
 * Cards index - export all Adaptive Card builders
 */

const approvalCard = require('./approvalCard');
const resultCard = require('./resultCard');
const pedmCard = require('./pedmCard');
const deviceCard = require('./deviceCard');
module.exports = {
  // Approval request cards
  buildRecordApprovalCard: approvalCard.buildRecordApprovalCard,
  buildRecordApprovalCardWithStatus: approvalCard.buildRecordApprovalCardWithStatus,
  buildRecordConfirmationCard: approvalCard.buildRecordConfirmationCard,
  buildRecordSearchResultsCard: approvalCard.buildRecordSearchResultsCard,
  buildRecordCreationCard: approvalCard.buildRecordCreationCard,
  buildCreateSecretSuccessCard: approvalCard.buildCreateSecretSuccessCard,
  buildPostCreateApprovalCard: approvalCard.buildPostCreateApprovalCard,
  buildPostCreateUidResolvingCard: approvalCard.buildPostCreateUidResolvingCard,
  buildPostCreateUidPendingCard: approvalCard.buildPostCreateUidPendingCard,
  buildRecordCreatedCard: approvalCard.buildRecordCreatedCard,
  buildRecordInvitationSentCard: approvalCard.buildRecordInvitationSentCard,
  buildRecordAlreadyHasAccessCard: approvalCard.buildRecordAlreadyHasAccessCard,
  buildRecordProcessingCard: approvalCard.buildRecordProcessingCard,
  buildFolderApprovalCard: approvalCard.buildFolderApprovalCard,
  buildFolderApprovalCardWithStatus: approvalCard.buildFolderApprovalCardWithStatus,
  buildFolderConfirmationCard: approvalCard.buildFolderConfirmationCard,
  buildFolderSearchResultsCard: approvalCard.buildFolderSearchResultsCard,
  buildFolderInvitationSentCard: approvalCard.buildFolderInvitationSentCard,
  buildFolderAlreadyHasAccessCard: approvalCard.buildFolderAlreadyHasAccessCard,
  buildFolderProcessingCard: approvalCard.buildFolderProcessingCard,
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
  
  // Helpers
  formatPermission: resultCard.formatPermission,
  formatDate: resultCard.formatDate,
};
