/**
 * Approval Handlers Index
 * 
 * Re-exports all approval handler functions for backwards compatibility.
 */

const { DURATION_MAP, parseDuration } = require('./helpers');
const { handleRecordApproval, handleRecordDenial } = require('./recordHandler');
const { handleFolderApproval, handleFolderDenial } = require('./folderHandler');
const { handleShareApproval, handleShareDenial } = require('./shareHandler');
const { routeApprovalAction, routeApprovalActionWithCardResponse } = require('./router');
const { 
  handleRefreshApprovalCard, 
  handleInlineLookup, 
  handleResetCard, 
  handleShowCreateForm, 
  handleSubmitCreateRecord, 
  handleCancelCreateForm 
} = require('./cardActions');

module.exports = {
  routeApprovalAction,
  routeApprovalActionWithCardResponse,
  handleRecordApproval,
  handleRecordDenial,
  handleFolderApproval,
  handleFolderDenial,
  handleShareApproval,
  handleShareDenial,
  handleRefreshApprovalCard,
  handleInlineLookup,
  handleResetCard,
  handleShowCreateForm,
  handleSubmitCreateRecord,
  handleCancelCreateForm,
  parseDuration,
  DURATION_MAP,
};
