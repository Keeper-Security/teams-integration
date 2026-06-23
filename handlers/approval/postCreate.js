/**
 * Post-create record UID resolution (async lookup after record-add / nsf-record-add).
 */

const keeperClient = require('../../services/keeperClient');
const cards = require('../../cards');
const { storeApprovalStatus, getApprovalStatus, createLogger } = require('../../services');
const { pinApprovalActivityId, tryUpdateApprovalCard } = require('./helpers');

const log = createLogger('PostCreate');

/**
 * @param {Object} params
 * @returns {Object}
 */
function buildPostCreateCardFromStatus(status, data = {}) {
  const base = {
    approvalId: status.approvalId || data.approvalId,
    requesterName: status.requesterName || data.requesterName,
    requesterId: status.requesterId || data.requesterId,
    requesterEmail: status.requesterEmail || data.requesterEmail,
    requesterAadObjectId: status.requesterAadObjectId || data.requesterAadObjectId,
    justification: status.justification || data.justification,
    identifier: status.identifier || data.identifier,
    recordTitle: status.recordTitle || data.recordTitle,
  };

  if (status.status === 'post_create_ready' && status.recordUid) {
    return cards.buildPostCreateApprovalCard({
      ...base,
      recordUid: status.recordUid,
      isNsf: !!status.isNsf,
      selfDestructEnabled: !!status.selfDestructEnabled,
      selfDestructDuration: status.selfDestructDuration || '24h',
    });
  }

  if (status.status === 'post_create_failed') {
    return cards.buildPostCreateUidPendingCard(base);
  }

  return cards.buildPostCreateUidResolvingCard({
    ...base,
    isNsf: !!status.isNsf,
  });
}

/**
 * Resolve a newly created record UID asynchronously and update the card in-place.
 * @param {Object} context - Teams turn context
 * @param {Object} params
 */
async function processPostCreateUidLookupAsync(context, params) {
  const {
    approvalId,
    recordTitle,
    isNsf = false,
    selfDestructEnabled = false,
    selfDestructDuration = '24h',
    requesterName,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    justification,
    identifier,
  } = params;

  const statusBase = {
    approvalId,
    recordTitle,
    isNsf,
    selfDestructEnabled,
    selfDestructDuration,
    requesterName,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    justification,
    identifier,
    type: 'post_create',
  };

  storeApprovalStatus(approvalId, {
    ...statusBase,
    status: 'post_create_resolving',
  });

  try {
    const recordUid = await keeperClient.lookupRecordUidAfterCreate(recordTitle, { isNsf });

    if (recordUid) {
      storeApprovalStatus(approvalId, {
        ...statusBase,
        status: 'post_create_ready',
        recordUid,
      });
      const finalCard = cards.buildPostCreateApprovalCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        recordUid,
        recordTitle,
        isNsf,
        selfDestructEnabled,
        selfDestructDuration,
      });
      await tryUpdateApprovalCard(approvalId, finalCard, context);
      log.info('Post-create UID resolved', { approvalId, recordUid, recordTitle });
      return;
    }

    storeApprovalStatus(approvalId, {
      ...statusBase,
      status: 'post_create_failed',
    });
    const fallbackCard = cards.buildPostCreateUidPendingCard({
      approvalId,
      requesterName,
      requesterId,
      requesterEmail,
      requesterAadObjectId,
      justification,
      identifier,
      recordTitle,
    });
    await tryUpdateApprovalCard(approvalId, fallbackCard, context);
    log.warn('Post-create UID lookup failed; showing manual search card', { approvalId, recordTitle });
  } catch (error) {
    log.error('Post-create UID lookup error', { approvalId, error: error.message });
    storeApprovalStatus(approvalId, {
      ...statusBase,
      status: 'post_create_failed',
    });
    try {
      await tryUpdateApprovalCard(approvalId, cards.buildPostCreateUidPendingCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        recordTitle,
      }), context);
    } catch (updateError) {
      log.debug('Failed to update post-create fallback card', updateError.message);
    }
  }
}

/**
 * Start async UID lookup after create; pins activity ID and returns resolving card.
 * @param {Object} context
 * @param {Object} params
 * @returns {Object} Adaptive Card
 */
function startPostCreateUidLookup(context, params) {
  pinApprovalActivityId(context, params.approvalId);
  processPostCreateUidLookupAsync(context, params).catch((err) => {
    log.error('Unhandled post-create UID lookup failure', { approvalId: params.approvalId, error: err.message });
  });
  return cards.buildPostCreateUidResolvingCard({
    approvalId: params.approvalId,
    requesterName: params.requesterName,
    requesterId: params.requesterId,
    requesterEmail: params.requesterEmail,
    requesterAadObjectId: params.requesterAadObjectId,
    justification: params.justification,
    identifier: params.identifier,
    recordTitle: params.recordTitle,
    isNsf: params.isNsf,
  });
}

/**
 * Refresh handler for post-create resolving cards.
 * @param {Object} data
 * @returns {Object|null}
 */
function handleRefreshPostCreateCard(data) {
  const { approvalId } = data;
  if (!approvalId) return null;

  const status = getApprovalStatus(approvalId);
  if (!status || status.type !== 'post_create') return null;

  if (status.status === 'post_create_resolving') {
    log.debug('Post-create still resolving', { approvalId });
    return null;
  }

  return buildPostCreateCardFromStatus(status, data);
}

module.exports = {
  startPostCreateUidLookup,
  handleRefreshPostCreateCard,
  processPostCreateUidLookupAsync,
};
