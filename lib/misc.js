'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.USER_AGENT = undefined;
exports.formatResponse = formatResponse;
exports.getAuthHeaders = getAuthHeaders;

var _lodash = require('lodash');

var USER_AGENT = exports.USER_AGENT = 'SafetyCulture SharePoint';

/**
* Formats a response to replace '_x0020_' with spaces.
* @param {object} res Response to deep replace
* @returns {object} Formatted response
*/
function formatResponse(res) {
  return (0, _lodash.transform)(res, function (result, val, key) {
    var newVal = val;
    var newKey = (0, _lodash.isString)(key) ? key.replace(/_x0020_/g, ' ') : key;

    if ((0, _lodash.isArray)(val)) newVal = (0, _lodash.map)(val, formatResponse);
    if ((0, _lodash.isObject)(val)) newVal = formatResponse(val);

    result[newKey] = newVal;
  });
}

/**
 * Support either token (oauth2) or cookie based authentication
 * to sharepoint API
 */
function getAuthHeaders(auth) {
  if (auth.token !== undefined) {
    return { 'Authorization': 'Bearer ' + auth.token };
  }

  return { 'Cookie': 'FedAuth=' + auth.FedAuth + ';rtFa=' + auth.rtFa + ';',
    'X-RequestDigest': auth.requestDigest };
}