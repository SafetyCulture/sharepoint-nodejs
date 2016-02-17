'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
var LIST_TEMPLATES = exports.LIST_TEMPLATES = {
  STANDARD: 100,
  LIBRARY: 101
};

// String builder for list URIS
var listURI = exports.listURI = function listURI(title) {
  return '/lists/GetByTitle(\'' + title + '\')';
};
var listType = exports.listType = function listType(name) {
  return 'SP.Data.' + name.replace(/ /g, '_x0020_') + 'Item';
};