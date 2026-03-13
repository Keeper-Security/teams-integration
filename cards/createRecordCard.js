/**
 * Create Record Card
 * Builds the Adaptive Card for creating a new record in the Task Module
 */

/**
 * Build the Create Record modal card
 * @param {Object} approvalContext - Context from the approval request
 * @param {string} searchQuery - Original search query (pre-fill title)
 * @returns {Object} - Adaptive Card for creating a new record
 */
function buildCreateRecordModal(approvalContext = {}, searchQuery = '') {
  const body = [
    {
      type: 'TextBlock',
      text: `Creating record for: **${approvalContext.requesterName || 'Requester'}**`,
      wrap: true,
    },
    {
      type: 'TextBlock',
      text: '_After creation, you\'ll be able to review and approve sharing_',
      wrap: true,
      isSubtle: true,
      size: 'Small',
    },
    
    // Divider
    {
      type: 'Container',
      style: 'emphasis',
      items: [],
      spacing: 'Medium',
    },
    
    // Title (Required) - pre-filled with search query
    {
      type: 'TextBlock',
      text: 'Title *',
      weight: 'Bolder',
    },
    {
      type: 'Input.Text',
      id: 'recordTitle',
      placeholder: 'Enter record title',
      value: searchQuery || '',
      isRequired: true,
      errorMessage: 'Title is required',
    },
    
    // Login (Required)
    {
      type: 'TextBlock',
      text: 'Login *',
      weight: 'Bolder',
    },
    {
      type: 'Input.Text',
      id: 'recordLogin',
      placeholder: 'Email or Username',
      isRequired: true,
      errorMessage: 'Login is required',
    },
    
    // Password
    {
      type: 'TextBlock',
      text: 'Password',
      weight: 'Bolder',
    },
    {
      type: 'Input.Text',
      id: 'recordPassword',
      placeholder: 'Enter password or use $GEN to auto-generate',
      style: 'password',
    },
    {
      type: 'TextBlock',
      text: 'Leave empty or enter $GEN to auto-generate a secure password',
      wrap: true,
      isSubtle: true,
      size: 'Small',
    },
    
    // Website URL (Optional)
    {
      type: 'TextBlock',
      text: 'Website URL',
      weight: 'Bolder',
    },
    {
      type: 'Input.Text',
      id: 'recordUrl',
      placeholder: 'https://',
    },
    
    // Notes (Optional)
    {
      type: 'TextBlock',
      text: 'Notes',
      weight: 'Bolder',
    },
    {
      type: 'Input.Text',
      id: 'recordNotes',
      placeholder: 'Additional notes',
      isMultiline: true,
    },
    
    // Hidden context field
    {
      type: 'Input.Text',
      id: 'approvalContext',
      isVisible: false,
      value: JSON.stringify(approvalContext),
    },
  ];

  const card = {
    type: 'AdaptiveCard',
    version: '1.2',
    body: body,
    actions: [
      {
        type: 'Action.Submit',
        title: 'Create Record',
        style: 'positive',
        data: { action: 'submit_create_record' },
      },
      {
        type: 'Action.Submit',
        title: 'Cancel',
        data: { action: 'cancel_create_record' },
      },
    ],
  };

  return card;
}

/**
 * Build a success card showing the newly created record
 * @param {Object} params - Parameters
 * @param {string} params.title - Record title
 * @param {string} params.recordUid - Record UID
 * @param {Object} params.approvalContext - Context from the approval request
 * @returns {Object} - Task module response with search modal pre-selecting new record
 */
function buildCreateRecordSuccessResponse(params) {
  const { title, recordUid, approvalContext } = params;
  
  return {
    task: {
      type: 'message',
      value: `Record "${title}" created successfully!\n\nUID: ${recordUid}\n\nPlease search for this record to approve access.`,
    },
  };
}

module.exports = {
  buildCreateRecordModal,
  buildCreateRecordSuccessResponse,
};
