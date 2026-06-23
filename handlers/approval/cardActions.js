/**
 * Card Actions Handler
 * 
 * Handles card-based actions like refresh, inline lookup, reset, and create form.
 */

const keeperClient = require('../../services/keeperClient');
const cards = require('../../cards');
const { getApprovalStatus, createLogger } = require('../../services');
const { sanitizeDisplayField, isValidEmail } = require('./helpers');

const log = createLogger('CardActions');

/**
 * Handle refresh action for approval cards
 * This is called when Teams auto-refreshes a card (via the refresh property)
 * Returns the appropriate card based on the current approval status
 * 
 * @param {Object} data - Refresh action data containing approvalId, type, and original card data
 * @returns {Object|null} - Updated card if status has changed, null otherwise
 */
async function handleRefreshApprovalCard(data) {
  const { approvalId, type } = data;
  
  if (!approvalId) {
    log.debug('Refresh: No approvalId provided');
    return null;
  }
  
  log.debug(`Checking refresh for approval ${approvalId} (type: ${type})`);
  
  const status = getApprovalStatus(approvalId);
  
  if (!status) {
    log.debug(`Approval ${approvalId} not yet processed, keeping original card`);
    return null;
  }
  
  log.debug(`Approval ${approvalId} has status: ${status.status}`);
  
  if (type === 'record') {
    if (status.status === 'invitation_sent') {
      return cards.buildRecordInvitationSentCard({
        approvalId,
        requesterName: status.requesterName || data.requesterName,
        requesterEmail: status.requesterEmail || data.requesterEmail,
        recordTitle: status.recordTitle || data.recordTitle,
        recordUid: status.recordUid || data.recordUid,
        justification: status.justification || data.justification,
        permission: status.permission,
        approverName: status.approverName,
        processedTime: status.processedTime,
      });
    }

    return cards.buildRecordApprovalCardWithStatus({
      approvalId: approvalId,
      requesterName: status.requesterName || data.requesterName,
      requesterEmail: status.requesterEmail || data.requesterEmail,
      recordTitle: status.recordTitle || data.recordTitle,
      justification: status.justification || data.justification,
      status: status.status,
      approverName: status.approverName,
      permission: status.permission,
      duration: status.duration,
      expiresAt: status.expiresAt,
      processedTime: status.processedTime,
      isNsf: status.isNsf,
      selfDestruct: status.selfDestruct,
      selfDestructDuration: status.selfDestructDuration,
    });
  } else if (type === 'folder') {
    if (status.status === 'invitation_sent') {
      return cards.buildFolderInvitationSentCard({
        approvalId,
        requesterName: status.requesterName || data.requesterName,
        requesterEmail: status.requesterEmail || data.requesterEmail,
        folderName: status.folderName || data.folderName,
        folderUid: status.folderUid || data.folderUid,
        justification: status.justification || data.justification,
        permission: status.permission,
        approverName: status.approverName,
        processedTime: status.processedTime,
      });
    }

    return cards.buildFolderApprovalCardWithStatus({
      approvalId: approvalId,
      requesterName: status.requesterName || data.requesterName,
      requesterEmail: status.requesterEmail || data.requesterEmail,
      folderName: status.folderName || data.folderName,
      justification: status.justification || data.justification,
      status: status.status,
      approverName: status.approverName,
      permission: status.permission,
      duration: status.duration,
      expiresAt: status.expiresAt,
      processedTime: status.processedTime,
      isNsf: status.isNsf,
    });
  }
  
  log.debug(`Unknown type ${type} for refresh`);
  return null;
}

/**
 * Handle inline lookup actions (search from the card itself)
 * Called when user clicks "Search" button on the approval card
 * Now supports multiple results with dropdown selection
 * 
 * @param {string} verb - 'lookup_record', 'lookup_folder', or 'lookup_share'
 * @param {Object} data - Card action data
 * @param {string} searchQuery - The search query from the input field
 * @returns {Object} - Updated card with search results
 */
async function handleInlineLookup(verb, data, searchQuery) {
  const isFolder = verb === 'lookup_folder';
  const isShare = verb === 'lookup_share';
  const {
    approvalId,
    identifier: rawIdentifier,
    recordTitle: rawRecordTitle,
    folderName: rawFolderName,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    requesterName: rawRequesterName,
    justification: rawJustification,
  } = data;

  if (requesterEmail && !isValidEmail(requesterEmail)) {
    log.warn('Invalid requesterEmail format rejected in lookup action', { approvalId });
    return null;
  }

  // Sanitize display-only fields so injected payloads cannot be
  // reflected back into card data with dangerous HTML characters.
  const requesterName = sanitizeDisplayField(rawRequesterName);
  const justification = sanitizeDisplayField(rawJustification);
  const recordTitle   = sanitizeDisplayField(rawRecordTitle);
  const folderName    = sanitizeDisplayField(rawFolderName);
  const identifier    = sanitizeDisplayField(rawIdentifier);

  const query = searchQuery || identifier || (isFolder ? folderName : recordTitle) || '';
  
  const lookupType = isFolder ? 'folder' : (isShare ? 'share' : 'record');
  log.debug(`Inline ${lookupType} lookup`, { query, approvalId });
  
  if (!query.trim()) {
    if (isFolder) {
      return cards.buildFolderSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        noResults: true,
        originalFolderName: folderName,
      });
    } else if (isShare) {
      return cards.buildShareSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        noResults: true,
        originalRecordTitle: recordTitle,
      });
    } else {
      return cards.buildRecordSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        noResults: true,
        originalRecordTitle: recordTitle,
      });
    }
  }
  
  try {
    const results = isFolder
      ? await keeperClient.searchFolders(query, 10)
      : await keeperClient.searchRecords(query, 10);
    
    log.debug(`Search results for "${query}": ${results?.length || 0}`);
    
    if (!results || results.length === 0) {
      if (isFolder) {
        return cards.buildFolderSearchResultsCard({
          approvalId,
          requesterName,
          requesterId,
          requesterEmail,
          requesterAadObjectId,
          justification,
          identifier,
          searchQuery: query,
          noResults: true,
          originalFolderName: folderName,
        });
      } else if (isShare) {
        return cards.buildShareSearchResultsCard({
          approvalId,
          requesterName,
          requesterId,
          requesterEmail,
          requesterAadObjectId,
          justification,
          identifier,
          searchQuery: query,
          noResults: true,
          originalRecordTitle: recordTitle,
        });
      } else {
        return cards.buildRecordSearchResultsCard({
          approvalId,
          requesterName,
          requesterId,
          requesterEmail,
          requesterAadObjectId,
          justification,
          identifier,
          searchQuery: query,
          noResults: true,
          originalRecordTitle: recordTitle,
        });
      }
    }
    
    if (isFolder) {
      const foundFolders = results.map(f => ({
        uid: f.uid || f.folder_uid,
        name: f.name || f.title || f.uid,
        isNsf: !!f.isNsf,
      }));

      const allNsf = foundFolders.length > 0 && foundFolders.every(f => f.isNsf);

      // PAM eligibility only applies to classic (non-NSF) folders.
      let isPamFolder = false;
      if (foundFolders.length > 0 && !allNsf) {
        try {
          isPamFolder = await keeperClient.isPamUserFolder(foundFolders[0].uid);
        } catch (e) {
          log.debug('isPamUserFolder check failed, defaulting to false', e.message);
        }
      }
      
      return cards.buildFolderSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        foundFolders,
        originalFolderName: folderName,
        isPamFolder,
        isNsf: allNsf,
      });
    } else if (isShare) {
      // One-Time Shares only work on classic records. Commander's
      // one-time-share supports neither PAM nor Nested Share Folder (NSF)
      // records, so filter both out of the OTS search results.
      const pamRecordTypes = ['pamdirectory', 'pamdatabase', 'pammachine', 'pamuser', 'pamremotebrowser'];
      const filteredResults = results.filter(r => {
        const recordType = (r.recordType || r.record_type || '').toLowerCase();
        const isPam = pamRecordTypes.some(pamType => recordType.includes(pamType));
        return !isPam && !r.isNsf;
      });
      
      if (filteredResults.length === 0) {
        log.debug('All search results were PAM/NSF records (ineligible for one-time share), showing no results message');
        return cards.buildShareSearchResultsCard({
          approvalId,
          requesterName,
          requesterId,
          requesterEmail,
          requesterAadObjectId,
          justification,
          identifier,
          searchQuery: query,
          noResults: true,
          pamRecordsOnly: true,
          originalRecordTitle: recordTitle,
        });
      }
      
      const foundRecords = filteredResults.map(r => ({
        uid: r.uid || r.record_uid,
        title: r.title || r.name || r.uid,
      }));
      
      return cards.buildShareSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        foundRecords,
        originalRecordTitle: recordTitle,
      });
    } else {
      const foundRecords = results.map(r => ({
        uid: r.uid || r.record_uid,
        title: r.title || r.name || r.uid,
        recordType: r.recordType || r.record_type || 'login',
        isNsf: !!r.isNsf,
      }));

      const allNsf = foundRecords.length > 0 && foundRecords.every(r => r.isNsf);
      
      return cards.buildRecordSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        foundRecords,
        originalRecordTitle: recordTitle,
        isNsf: allNsf,
      });
    }
  } catch (error) {
    log.error(`Error searching ${lookupType}s`, error);
    
    if (isFolder) {
      return cards.buildFolderSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        noResults: true,
        originalFolderName: folderName,
      });
    } else if (isShare) {
      return cards.buildShareSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        noResults: true,
        originalRecordTitle: recordTitle,
      });
    } else {
      return cards.buildRecordSearchResultsCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        searchQuery: query,
        noResults: true,
        originalRecordTitle: recordTitle,
      });
    }
  }
}

/**
 * Handle reset card actions - returns the original approval card
 * Called when user clicks "Reset" button to go back to the initial search state
 * 
 * @param {string} verb - 'reset_record_card', 'reset_folder_card', or 'reset_share_card'
 * @param {Object} data - Card action data
 * @returns {Object} - Original approval card
 */
function handleResetCard(verb, data) {
  const isFolder = verb === 'reset_folder_card';
  const isShare = verb === 'reset_share_card';
  const {
    approvalId,
    identifier,
    recordTitle,
    folderName,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    requesterName,
    justification,
  } = data;
  
  const resetType = isFolder ? 'folder' : (isShare ? 'share' : 'record');
  log.debug(`Resetting ${resetType} card`, { approvalId });
  
  if (isFolder) {
    return cards.buildFolderApprovalCard({
      approvalId,
      requesterName,
      requesterId,
      requesterEmail,
      requesterAadObjectId,
      folderName: folderName,
      folderUid: null,
      justification,
      isUid: false,
      identifier: identifier || folderName,
    });
  } else if (isShare) {
    return cards.buildOneTimeShareApprovalCard({
      approvalId,
      requesterName,
      requesterId,
      requesterEmail,
      requesterAadObjectId,
      recordTitle: recordTitle,
      recordUid: null,
      justification,
      isUid: false,
      identifier: identifier || recordTitle,
    });
  } else {
    return cards.buildRecordApprovalCard({
      approvalId,
      requesterName,
      requesterId,
      requesterEmail,
      requesterAadObjectId,
      recordTitle: recordTitle,
      recordUid: null,
      justification,
      isUid: false,
      identifier: identifier || recordTitle,
    });
  }
}

/**
 * Handle show_create_form action - returns inline create record form card
 * 
 * @param {Object} data - Card action data
 * @returns {Object} - Create record form card
 */
function handleShowCreateForm(data) {
  const {
    approvalId,
    identifier,
    recordTitle,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    requesterName,
    justification,
    searchQuery,
  } = data;
  
  log.debug('Showing create record form', { approvalId });
  
  return cards.buildRecordCreationCard({
    approvalId,
    requesterName,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    justification,
    identifier,
    originalRecordTitle: recordTitle,
    searchQuery: searchQuery || recordTitle,
  });
}

/**
 * Handle submit_create_record action - creates record and returns card with approval options
 * 
 * @param {Object} data - Card action data (approval context)
 * @param {Object} formData - Form input data
 * @returns {Object} - Created record card with approval options, or error card
 */
async function handleSubmitCreateRecord(data, formData) {
  const {
    approvalId,
    identifier,
    originalRecordTitle,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    requesterName,
    justification,
  } = data;
  
  const {
    recordTitle: rawRecordTitle,
    recordLogin: rawRecordLogin,
    recordPassword,
    recordUrl: rawRecordUrl,
    recordNotes,
  } = formData;

  const recordTitle = (rawRecordTitle || '').replace(/[\r\n\x00-\x1f\x7f]/g, '');
  const recordLogin = (rawRecordLogin || '').replace(/[\r\n\x00-\x1f\x7f]/g, '');
  const recordUrl = (rawRecordUrl || '').replace(/[\r\n\x00-\x1f\x7f]/g, '');

  log.debug('Creating record', { title: recordTitle, approvalId });
  
  if (!recordTitle || !recordTitle.trim()) {
    return cards.buildRecordCreationCard({
      approvalId,
      requesterName,
      requesterId,
      requesterEmail,
      requesterAadObjectId,
      justification,
      identifier,
      originalRecordTitle,
      searchQuery: '',
      error: 'Title is required',
    });
  }
  
  if (!recordLogin || !recordLogin.trim()) {
    return cards.buildRecordCreationCard({
      approvalId,
      requesterName,
      requesterId,
      requesterEmail,
      requesterAadObjectId,
      justification,
      identifier,
      originalRecordTitle,
      searchQuery: recordTitle,
      error: 'Login is required',
    });
  }

  const trimmedUrl = recordUrl?.trim() || '';
  if (trimmedUrl) {
    let urlValid = false;
    try {
      const parsed = new URL(trimmedUrl);
      urlValid = parsed.protocol === 'http:' || parsed.protocol === 'https:';
    } catch {
      urlValid = false;
    }
    if (!urlValid) {
      return cards.buildRecordCreationCard({
        approvalId,
        requesterName,
        requesterId,
        requesterEmail,
        requesterAadObjectId,
        justification,
        identifier,
        originalRecordTitle,
        searchQuery: recordTitle,
        error: 'URL must start with http:// or https://',
      });
    }
  }

  const generatePassword = !recordPassword || recordPassword.trim() === '' || recordPassword === '$GEN';
  const passwordToUse = generatePassword ? '$GEN' : recordPassword;
  
  const result = await keeperClient.createRecord({
    title: recordTitle.trim(),
    login: recordLogin.trim(),
    password: passwordToUse,
    url: recordUrl?.trim() || null,
    notes: recordNotes?.trim() || null,
    generatePassword: generatePassword,
  });
  
  if (!result.success) {
    log.error('Failed to create record', { error: result.error, errorCode: result.errorCode });

    // Re-render the same form with a friendly inline error so the user keeps
    // their typed values and can fix the offending field in-place.
    // On password-complexity failures, blank the password input so the user
    // re-enters or leaves it empty to use the auto-generator ($GEN).
    const isPasswordPolicyError = result.errorCode === 'POLICY_PASSWORD_COMPLEXITY';
    const errorTitle = result.errorTitle || 'Record creation failed';
    const errorBody = result.error || 'Unknown error';
    const tip = isPasswordPolicyError
      ? '\n\nTip: leave the password empty or use the auto-generate option to get a Keeper-compliant password.'
      : '';

    return cards.buildRecordCreationCard({
      approvalId,
      requesterName,
      requesterId,
      requesterEmail,
      requesterAadObjectId,
      justification,
      identifier,
      originalRecordTitle,
      searchQuery: recordTitle,
      createSecretFlow: false,
      recordTitle,
      recordLogin,
      recordPassword: isPasswordPolicyError ? '' : recordPassword,
      recordUrl,
      recordNotes,
      error: `${errorTitle}\n${errorBody}${tip}`,
    });
  }
  
  log.debug('Record created successfully', result.recordUid);
  
  const safeApprovalId = approvalId || '';
  const safeRecordUid = result.recordUid;
  const safeRecordTitle = recordTitle.trim();
  const safeRequesterId = requesterId || '';
  const safeRequesterEmail = requesterEmail || '';
  const safeRequesterName = requesterName || 'Unknown';
  const safeJustification = justification || '';
  
  const createdRecordCard = {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      { type: 'TextBlock', text: 'Record Created Successfully!', weight: 'Bolder', size: 'Large' },
      { type: 'TextBlock', text: `Requester: ${safeRequesterName}`, wrap: true },
      { type: 'TextBlock', text: `Record: ${safeRecordTitle}`, wrap: true, weight: 'Bolder' },
      { type: 'TextBlock', text: `UID: ${safeRecordUid}`, wrap: true, size: 'Small' },
    ],
    actions: [
      {
        type: 'Action.Execute',
        title: 'Approve (View Only, 1h)',
        style: 'positive',
        verb: 'approve_record',
        data: { 
          action: 'approve_record', 
          approvalId: safeApprovalId, 
          recordUid: safeRecordUid, 
          recordTitle: safeRecordTitle, 
          requesterId: safeRequesterId, 
          requesterEmail: safeRequesterEmail, 
          requesterName: safeRequesterName, 
          justification: safeJustification,
          permission: 'view_only',
          duration: '1h',
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny',
        style: 'destructive',
        verb: 'deny_record',
        data: { 
          action: 'deny_record', 
          approvalId: safeApprovalId, 
          recordUid: safeRecordUid, 
          recordTitle: safeRecordTitle, 
          requesterId: safeRequesterId, 
          requesterEmail: safeRequesterEmail, 
          requesterName: safeRequesterName, 
          justification: safeJustification,
        },
      },
    ],
  };
  
  log.debug('Returning created record card');
  return createdRecordCard;
}

/**
 * Handle cancel_create_form action - returns to search results card with re-fetched results
 * 
 * @param {Object} data - Card action data
 * @returns {Object} - Search results card with actual search results
 */
async function handleCancelCreateForm(data) {
  const {
    approvalId,
    identifier,
    recordTitle,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    requesterName,
    justification,
    searchQuery,
  } = data;
  
  const query = searchQuery || recordTitle || identifier;
  log.debug('Cancelling create form, re-running search', { approvalId, query });
  
  let searchResults = [];
  let noResults = true;
  
  if (query) {
    try {
      searchResults = await keeperClient.searchRecords(query);
      noResults = !searchResults || searchResults.length === 0;
      log.debug(`Search results for "${query}": ${searchResults.length}`);
    } catch (error) {
      log.error('Error re-running search', error);
    }
  }
  
  return cards.buildRecordSearchResultsCard({
    approvalId,
    requesterName,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    justification,
    identifier,
    searchQuery: query,
    foundRecords: searchResults,
    noResults,
    originalRecordTitle: recordTitle,
  });
}

module.exports = {
  handleRefreshApprovalCard,
  handleInlineLookup,
  handleResetCard,
  handleShowCreateForm,
  handleSubmitCreateRecord,
  handleCancelCreateForm,
};
