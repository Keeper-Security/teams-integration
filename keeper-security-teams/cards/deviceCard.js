/**
 * Adaptive Card builders for Cloud SSO Device Approval
 * 
 * These cards are displayed when device approval requests are detected,
 * allowing admins to approve or deny new device registrations.
 */

/**
 * Build an Adaptive Card for device approval request
 */
function buildDeviceApprovalCard({
  deviceId,
  deviceName,
  deviceType,
  username,
  email,
  ipAddress,
  location,
  created,
}) {
  const icon = getDeviceIcon(deviceType);
  
  const facts = [
    { title: 'User', value: username || email || 'Unknown' },
  ];
  
  if (email && email !== username) {
    facts.push({ title: 'Email', value: email });
  }
  
  if (deviceName) {
    facts.push({ title: 'Device', value: deviceName });
  }
  
  if (deviceType) {
    facts.push({ title: 'Type', value: deviceType });
  }
  
  if (ipAddress) {
    facts.push({ title: 'IP Address', value: ipAddress });
  }
  
  if (location) {
    facts.push({ title: 'Location', value: location });
  }
  
  if (created) {
    facts.push({ title: 'Requested', value: formatDate(created) });
  }
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: 'emphasis',
        items: [
          {
            type: 'ColumnSet',
            columns: [
              {
                type: 'Column',
                width: 'auto',
                items: [{ type: 'TextBlock', text: icon, size: 'ExtraLarge' }],
              },
              {
                type: 'Column',
                width: 'stretch',
                items: [
                  { type: 'TextBlock', text: 'Device Approval Request', weight: 'Bolder', size: 'Large', color: 'Accent' },
                  { type: 'TextBlock', text: 'ID: ' + deviceId, size: 'Small', isSubtle: true },
                ],
              },
            ],
          },
        ],
      },
      {
        type: 'Container',
        items: [
          { type: 'FactSet', facts: facts },
        ],
      },
      {
        type: 'Container',
        items: [
          { 
            type: 'TextBlock', 
            text: '⚠️ Only approve this device if you recognize the user and expect them to be logging in from this device.', 
            wrap: true, 
            size: 'Small', 
            isSubtle: true 
          },
        ],
      },
    ],
    actions: [
      {
        type: 'Action.Submit',
        title: '✅ Approve',
        style: 'positive',
        data: {
          action: 'approve_device',
          deviceId: deviceId,
          deviceName: deviceName,
          username: username || email,
        },
      },
      {
        type: 'Action.Submit',
        title: '❌ Deny',
        style: 'destructive',
        data: {
          action: 'deny_device',
          deviceId: deviceId,
          deviceName: deviceName,
          username: username || email,
        },
      },
    ],
  };
}

/**
 * Build an Adaptive Card showing device was approved
 */
function buildDeviceApprovedCard(approverName, deviceName, username) {
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: 'good',
        items: [
          {
            type: 'ColumnSet',
            columns: [
              {
                type: 'Column',
                width: 'auto',
                items: [{ type: 'TextBlock', text: '✅', size: 'Large' }],
              },
              {
                type: 'Column',
                width: 'stretch',
                items: [
                  { type: 'TextBlock', text: 'Device Approved', weight: 'Bolder', size: 'Medium', color: 'Good' },
                ],
              },
            ],
          },
        ],
      },
      {
        type: 'Container',
        items: [
          {
            type: 'FactSet',
            facts: [
              { title: 'User', value: username },
              ...(deviceName ? [{ title: 'Device', value: deviceName }] : []),
              { title: 'Approved by', value: approverName },
            ],
          },
        ],
      },
    ],
  };
}

/**
 * Build an Adaptive Card showing device was denied
 */
function buildDeviceDeniedCard(approverName, deviceName, username) {
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: 'attention',
        items: [
          {
            type: 'ColumnSet',
            columns: [
              {
                type: 'Column',
                width: 'auto',
                items: [{ type: 'TextBlock', text: '❌', size: 'Large' }],
              },
              {
                type: 'Column',
                width: 'stretch',
                items: [
                  { type: 'TextBlock', text: 'Device Denied', weight: 'Bolder', size: 'Medium', color: 'Attention' },
                ],
              },
            ],
          },
        ],
      },
      {
        type: 'Container',
        items: [
          {
            type: 'FactSet',
            facts: [
              { title: 'User', value: username },
              ...(deviceName ? [{ title: 'Device', value: deviceName }] : []),
              { title: 'Denied by', value: approverName },
            ],
          },
        ],
      },
    ],
  };
}

function getDeviceIcon(deviceType) {
  const type = (deviceType || '').toLowerCase();
  if (type.includes('mobile') || type.includes('phone') || type.includes('ios') || type.includes('android')) {
    return '📱';
  }
  if (type.includes('tablet') || type.includes('ipad')) {
    return '📱';
  }
  if (type.includes('mac') || type.includes('windows') || type.includes('linux') || type.includes('desktop')) {
    return '💻';
  }
  if (type.includes('browser') || type.includes('web')) {
    return '🌐';
  }
  return '📱';
}

function formatDate(dateStr) {
  try {
    const date = new Date(dateStr);
    return date.toLocaleString();
  } catch (e) {
    return dateStr;
  }
}

module.exports = {
  buildDeviceApprovalCard,
  buildDeviceApprovedCard,
  buildDeviceDeniedCard,
};
