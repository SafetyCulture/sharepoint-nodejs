'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.fillSpaces = exports.listType = exports.listURI = exports.LIST_TEMPLATES = undefined;

var _lodash = require('lodash');

var _lodash2 = _interopRequireDefault(_lodash);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var LIST_TEMPLATES = exports.LIST_TEMPLATES = {
  STANDARD: 100,
  LIBRARY: 101
};

// String builder for list URIS
var listURI = exports.listURI = function listURI(title) {
  return '/lists/GetByTitle(\'' + title + '\')';
};
var listType = exports.listType = function listType(name) {
  return 'SP.Data.' + name.replace(/\/|\-/g, '').replace(/ /g, '_x0020_') + 'Item';
};

// Small helper to replace spaces in keys with '_x0020_' within an object
var fillSpaces = exports.fillSpaces = function fillSpaces(data) {
  return _lodash2.default.mapKeys(data, function (val, key) {
    return key.replace(/ /g, '_x0020_');
  });
};