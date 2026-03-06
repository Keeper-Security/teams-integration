/**
 * Background services index
 */

const EpmPoller = require('./pedmPoller');
const DevicePoller = require('./devicePoller');

module.exports = {
  EpmPoller,
  DevicePoller,
};
