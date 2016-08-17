'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.fillSpaces = exports.libraryType = exports.listType = exports.sharepointEscapeChars = exports.listURI = exports.LIST_TEMPLATES = exports.AFFIXED_NEWLINE_REGEX = exports.STD_NEWLINE_REGEX = undefined;

var _lodash = require('lodash');

var STD_NEWLINE_REGEX = exports.STD_NEWLINE_REGEX = /(\n|\r|\\r|\\n)/gm;
var AFFIXED_NEWLINE_REGEX = exports.AFFIXED_NEWLINE_REGEX = /^(\n|\r|\\r|\\n)$/gm;
var LIST_TEMPLATES = exports.LIST_TEMPLATES = {
  STANDARD: 100,
  LIBRARY: 101
};

// String builder for list URIS
// Title must be URI Encoded and Apostrophes must be duplicated to avoid errors
var listURI = exports.listURI = function listURI(title) {
  return '/lists/GetByTitle(\'' + encodeURI(title).replace(/\'/g, "''") + '\')';
};

// This function converts a string to the encoding style Sharepoint uses.
// 1. Handles newlines and newline strings.
// 2. URI/Percent encode the string.
// 3. Using regex, replace all encoded symbol codes with the Sharepoint equivalent.
// 4. Finish off by explicitly replacing safe URI symbols with Sharepoint codes.
// References:
// http://www.blooberry.com/indexdot/html/topics/urlencoding.htm
// https://abstractspaces.wordpress.com/2008/05/07/sharepoint-column-names-internal-name-mappings-for-non-alphabet/
var sharepointEscapeChars = exports.sharepointEscapeChars = function sharepointEscapeChars(str) {
  // Strip leading/trailing newlines and replace others with whitespace
  var result = str.replace(AFFIXED_NEWLINE_REGEX, '').replace(STD_NEWLINE_REGEX, ' ');
  return encodeURI(result).replace(/(\%)([a-zA-Z0-9]{2})/g, '_x00$2_') // convert all unreserved
  .replace(/\$/g, '_x0024_') // $
  .replace(/\-/g, '_x002d_') // -
  .replace(/\./g, '_x002e_') // .
  .replace(/\+/g, '_x002b_') // +
  .replace(/\!/g, '_x0021_') // !
  .replace(/\&/g, '_x0026_') // &
  .replace(/\)/g, '_x0029_') // (
  .replace(/\(/g, '_x0028_') // )
  .replace(/\?/g, '_x003f_') // ?
  .replace(/\=/g, '_x003d_') // =
  .replace(/\*/g, '_x002a_') // *
  .replace(/\,/g, '_x002c_') // ,
  .replace(/\//g, '_x002f_') // /
  .replace(/\'/g, '_x0027_') // '
  .replace(/\@/g, '_x0040_') // @
  .replace(/\:/g, '_x003a_') // :
  .replace(/\;/g, '_x003b_') // ;
  .replace(/\#/g, '_x0023_'); // #
};

var listType = exports.listType = function listType(name) {
  return 'SP.Data.' + sharepointEscapeChars(name.charAt(0).toUpperCase() + name.slice(1)) + 'ListItem';
};
var libraryType = exports.libraryType = function libraryType(name) {
  return 'SP.Data.' + sharepointEscapeChars(name.charAt(0).toUpperCase() + name.slice(1)) + 'Item';
};
// Small helper to replace spaces in keys with '_x0020_' within an object
var fillSpaces = exports.fillSpaces = function fillSpaces(data) {
  return (0, _lodash.mapKeys)(data, function (val, key) {
    return sharepointEscapeChars(key);
  });
};