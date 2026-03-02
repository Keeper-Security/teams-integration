/**
 * Services index - export all services
 */

const keeperClient = require('./keeperClient');
const { KeeperClient } = require('./keeperClient');
const graphService = require('./graphService');
const { createLogger } = require('./logger');
const ksmService = require('./ksmService');
const {
  ChannelService,
  initializeChannelService,
  getChannelService,
  storeConversationReference,
  getConversationReference,
  isApprovalsChannel,
  isTeamsChannel,
  storeApprovalActivityId,
  getApprovalActivityId,
  removeApprovalActivityId,
  storeApprovalStatus,
  getApprovalStatus,
  isApprovalProcessed,
} = require('./channelService');

module.exports = {
  keeperClient,
  KeeperClient,
  graphService,
  createLogger,
  ksmService,
  ChannelService,
  initializeChannelService,
  getChannelService,
  storeConversationReference,
  getConversationReference,
  isApprovalsChannel,
  isTeamsChannel,
  storeApprovalActivityId,
  getApprovalActivityId,
  removeApprovalActivityId,
  storeApprovalStatus,
  getApprovalStatus,
  isApprovalProcessed,
};
