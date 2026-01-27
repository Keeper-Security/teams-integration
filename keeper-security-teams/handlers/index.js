/**
 * Handlers index - export all handlers
 */

const commandHandler = require('./commandHandler');
const approvalHandler = require('./approvalHandler');
const pedmHandler = require('./pedmHandler');
const deviceHandler = require('./deviceHandler');

module.exports = {
  // Command handler
  routeCommand: commandHandler.routeCommand,
  handleRequestRecord: commandHandler.handleRequestRecord,
  handleRequestFolder: commandHandler.handleRequestFolder,
  handleShare: commandHandler.handleShare,
  handleSearch: commandHandler.handleSearch,
  handleHelp: commandHandler.handleHelp,
  handleStatus: commandHandler.handleStatus,
  parseCommand: commandHandler.parseCommand,
  generateApprovalId: commandHandler.generateApprovalId,
  
  // Approval handler
  routeApprovalAction: approvalHandler.routeApprovalAction,
  handleRecordApproval: approvalHandler.handleRecordApproval,
  handleRecordDenial: approvalHandler.handleRecordDenial,
  handleFolderApproval: approvalHandler.handleFolderApproval,
  handleFolderDenial: approvalHandler.handleFolderDenial,
  handleShareApproval: approvalHandler.handleShareApproval,
  handleShareDenial: approvalHandler.handleShareDenial,
  handleSearchRecordsAction: approvalHandler.handleSearchRecordsAction,
  parseDuration: approvalHandler.parseDuration,
  
  // PEDM handler
  routePedmAction: pedmHandler.routePedmAction,
  handlePedmApproval: pedmHandler.handlePedmApproval,
  handlePedmDenial: pedmHandler.handlePedmDenial,
  
  // Device handler
  routeDeviceAction: deviceHandler.routeDeviceAction,
  handleDeviceApproval: deviceHandler.handleDeviceApproval,
  handleDeviceDenial: deviceHandler.handleDeviceDenial,
};
