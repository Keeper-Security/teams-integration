/**
 * Card Actions Handler
 * 
 * Handles card-based actions like refresh, inline lookup, reset, and create form.
 */

const keeperClient = require('../../services/keeperClient');
const cards = require('../../cards');
const { getApprovalStatus, createLogger } = require('../../services');

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
    });
  } else if (type === 'folder') {
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
    identifier,
    recordTitle,
    folderName,
    requesterId,
    requesterEmail,
    requesterAadObjectId,
    requesterName,
    justification,
  } = data;
  
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
      }));
      
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
      });
    } else if (isShare) {
      const pamRecordTypes = ['pamdirectory', 'pamdatabase', 'pammachine', 'pamuser', 'pamremotebrowser'];
      const filteredResults = results.filter(r => {
        const recordType = (r.recordType || r.record_type || '').toLowerCase();
        return !pamRecordTypes.some(pamType => recordType.includes(pamType));
      });
      
      if (filteredResults.length === 0) {
        log.debug('All search results were PAM records, showing no results message');
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
      }));
      
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
  
  const { recordTitle, recordLogin, recordPassword, recordUrl, recordNotes } = formData;
  
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
    log.error('Failed to create record', result.error);
    return {
      type: 'AdaptiveCard',
      '$schema': 'https://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.4',
      body: [
        { type: 'TextBlock', text: 'Record Creation Failed', weight: 'Bolder', size: 'Large', color: 'Attention' },
        { type: 'TextBlock', text: `Error: ${result.error || 'Unknown error'}`, wrap: true },
        { type: 'TextBlock', text: 'Please try again or contact support.', wrap: true, isSubtle: true },
      ],
      actions: [
        {
          type: 'Action.Execute',
          title: 'Try Again',
          verb: 'show_create_form',
          data: { action: 'show_create_form', approvalId: approvalId || '', identifier: identifier || '', recordTitle: originalRecordTitle || '', requesterId: requesterId || '', requesterEmail: requesterEmail || '', requesterAadObjectId: requesterAadObjectId || '', requesterName: requesterName || '', justification: justification || '' },
        },
        {
          type: 'Action.Execute',
          title: 'Cancel',
          verb: 'cancel_create_form',
          data: { action: 'cancel_create_form', approvalId: approvalId || '', identifier: identifier || '', recordTitle: originalRecordTitle || '', requesterId: requesterId || '', requesterEmail: requesterEmail || '', requesterAadObjectId: requesterAadObjectId || '', requesterName: requesterName || '', justification: justification || '' },
        },
      ],
    };
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
    version: '1.4',
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
    records: searchResults,
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
