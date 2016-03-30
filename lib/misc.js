'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.USER_AGENT = undefined;
exports.formatResponse = formatResponse;

var _lodash = require('lodash');

var _lodash2 = _interopRequireDefault(_lodash);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var USER_AGENT = exports.USER_AGENT = 'SafetyCulture SharePoint';

/**
* Formats a response to replace '_x0020_' with spaces.
* @param {object} res Response to deep replace
* @returns {object} Formatted response
*/
function formatResponse(res) {
  return _lodash2.default.transform(res, function (result, val, key) {
    var newVal = val;
    var newKey = _lodash2.default.isString(key) ? key.replace(/_x0020_/g, ' ') : key;

    if (_lodash2.default.isArray(val)) newVal = _lodash2.default.map(val, formatResponse);
    if (_lodash2.default.isObject(val)) newVal = formatResponse(val);

    result[newKey] = newVal;
  });
}