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
          { type: 'TextBlock', text: 'Device Approval Request', weight: 'Bolder', size: 'Large', color: 'Accent' },
          { type: 'TextBlock', text: 'ID: ' + deviceId, size: 'Small', isSubtle: true },
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
            text: 'Only approve this device if you recognize the user and expect them to be logging in from this device.', 
            wrap: true, 
            size: 'Small', 
            isSubtle: true 
          },
        ],
      },
    ],
    actions: [
      {
        type: 'Action.Execute',
        title: 'Approve',
        style: 'positive',
        verb: 'approve_device',
        data: {
          action: 'approve_device',
          deviceId: deviceId,
          deviceName: deviceName,
          username: username || email,
        },
      },
      {
        type: 'Action.Execute',
        title: 'Deny',
        style: 'destructive',
        verb: 'deny_device',
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
 * Updated to include deviceId (matching EPM pattern)
 */
function buildDeviceApprovedCard(approverName, deviceName, username, deviceId) {
  const facts = [
    { title: 'User', value: username },
    { title: 'Approved by', value: approverName },
  ];
  
  if (deviceId) {
    facts.push({ title: 'Device ID', value: deviceId });
  }
  
  if (deviceName) {
    facts.push({ title: 'Device', value: deviceName });
  }
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: 'good',
        items: [
          { type: 'TextBlock', text: 'Device Approved', weight: 'Bolder', size: 'Medium', color: 'Good' },
        ],
      },
      {
        type: 'Container',
        items: [
          { type: 'FactSet', facts: facts },
        ],
      },
    ],
  };
}

/**
 * Build an Adaptive Card showing device was denied
 * Updated to include deviceId (matching EPM pattern)
 */
function buildDeviceDeniedCard(approverName, deviceName, username, deviceId) {
  const facts = [
    { title: 'User', value: username },
    { title: 'Denied by', value: approverName },
  ];
  
  if (deviceId) {
    facts.push({ title: 'Device ID', value: deviceId });
  }
  
  if (deviceName) {
    facts.push({ title: 'Device', value: deviceName });
  }
  
  return {
    type: 'AdaptiveCard',
    '$schema': 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      {
        type: 'Container',
        style: 'attention',
        items: [
          { type: 'TextBlock', text: 'Device Denied', weight: 'Bolder', size: 'Medium', color: 'Attention' },
        ],
      },
      {
        type: 'Container',
        items: [
          { type: 'FactSet', facts: facts },
        ],
      },
    ],
  };
}

function formatDate(dateStr) {
  try {
    const date = new Date(dateStr);
    return date.toLocaleString('en-US', {
      month: '2-digit',
      day: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true,
    });
  } catch (e) {
    return dateStr;
  }
}

module.exports = {
  buildDeviceApprovalCard,
  buildDeviceApprovedCard,
  buildDeviceDeniedCard,
};
