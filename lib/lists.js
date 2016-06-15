'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.fillSpaces = exports.libraryType = exports.listType = exports.sharepointEscapeChars = exports.listURI = exports.LIST_TEMPLATES = undefined;

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

// Escapes non-alphanumerical chars to appropriate SharePoint internal code
// https://abstractspaces.wordpress.com/2008/05/07/sharepoint-column-names-internal-name-mappings-for-non-alphabet/
var sharepointEscapeChars = exports.sharepointEscapeChars = function sharepointEscapeChars(str) {
  return str.replace(/ /g, '_x0020_') // whitespace
  .replace(/\`/g, '_x0060_') // backtick
  .replace(/\//g, '_x002f_') // forwardslash
  .replace(/\./g, '_x002e_') // period
  .replace(/\,/g, '_x002c_') // comma
  .replace(/\?/g, '_x003f_') // questionmark
  .replace(/\>/g, '_x003e_') // right angle bracket
  .replace(/\</g, '_x003c_') // left angle bracket
  .replace(/\\/g, '_x005c_') // backslash
  .replace(/\'/g, '_x0027_') // apostrophe
  .replace(/\;/g, '_x003b_') // semicolon
  .replace(/\|/g, '_x007c_') // pipe
  .replace(/\"/g, '_x0022_') // quotation
  .replace(/\:/g, '_x003a_') // colon
  .replace(/\}/g, '_x007d_') // right curly brace
  .replace(/\{/g, '_x007b_') // left curly brace
  .replace(/\=/g, '_x003d_') // equals sign
  .replace(/\-/g, '_x002d_') // minus sign
  .replace(/\+/g, '_x002b_') // plus sign
  .replace(/\)/g, '_x0029_') // right paranthesis
  .replace(/\(/g, '_x0028_') // left paranthesis
  .replace(/\*/g, '_x002a_') // asterisk
  .replace(/\&/g, '_x0026_') // ampersand
  .replace(/\^/g, '_x005e_') // caret
  .replace(/\%/g, '_x0025_') // percent
  .replace(/\$/g, '_x0024_') // dollar
  .replace(/\#/g, '_x0023_') // hash
  .replace(/\@/g, '_x0040_') // at symbol
  .replace(/\!/g, '_x0021_') // exclamation
  .replace(/\~/g, '_x007e_'); // tilde
};
var listType = exports.listType = function listType(name) {
  return 'SP.Data.' + sharepointEscapeChars(name.charAt(0).toUpperCase() + name.slice(1)) + 'ListItem';
};
var libraryType = exports.libraryType = function libraryType(name) {
  return 'SP.Data.' + sharepointEscapeChars(name.charAt(0).toUpperCase() + name.slice(1)) + 'Item';
};
// Small helper to replace spaces in keys with '_x0020_' within an object
var fillSpaces = exports.fillSpaces = function fillSpaces(data) {
  return _lodash2.default.mapKeys(data, function (val, key) {
    return sharepointEscapeChars(key);
  });
};